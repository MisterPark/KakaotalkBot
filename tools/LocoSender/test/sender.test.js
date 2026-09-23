'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { EventEmitter } = require('node:events');
const { spawn } = require('node:child_process');
const { Long, talk } = require('node-kakao');
const { buildMessage, preview, id } = require('../message');
const { LocoSender } = require('../sender');
const { loadProfile } = require('../profile');

const roomId = '9007199254740993';
const firstId = '9223372036854775807';
const secondId = '9007199254740995';
const mention = (userId = firstId, nickname = '동명이인') => ({ kind: 'mention', userId, nickname });
const text = value => ({ kind: 'text', text: value });

test('64비트 ID·동명이인·중복 멘션과 UTF-16 표시 길이', () => {
    const result = preview(roomId, [mention(), text(' '), mention(secondId), text(' '), mention(), text(' '), mention('12', '봇😀')]);
    assert.equal(result.chatId, roomId);
    assert.deepEqual(result.mentions, [
        { userId: firstId, at: [1, 3], length: 4 },
        { userId: secondId, at: [2], length: 4 },
        { userId: '12', at: [4], length: 3 }
    ]);
    assert.ok(result.extra.includes('"user_id":9223372036854775807'));
    assert.ok(!result.extra.includes('"user_id":"'));
});

test('부정확한 숫자·범위 초과·모호한 @ 거절', () => {
    for (const bad of [9007199254740993, 1, null, '0', '-1', '01', '1.2', '1e3', '9223372036854775808']) {
        assert.throws(() => id(bad, 'id'), { code: 'INVALID_ID' });
    }
    assert.throws(() => preview(roomId, [mention(), text('a@b')]), { code: 'AMBIGUOUS_AT' });
    assert.throws(() => preview(roomId, [mention(firstId, '')]), { code: 'INVALID_NICKNAME' });
    assert.throws(() => preview(roomId, [text('x'.repeat(4001))]), { code: 'INVALID_TEXT' });
    assert.equal(preview(roomId, [text('a@b')]).text, 'a@b');
});

test('실제 라이브러리 WRITE 직렬화 (네트워크를 가짜 세션으로 교체)', async () => {
    const { channelId, chat } = buildMessage(roomId, [mention(), text(' 확인')]);
    let sent;
    const session = new talk.TalkChannelSession({ channelId }, {
        clientUser: { userId: Long.fromString('11') },
        request: async (method, payload) => {
            sent = { method, payload };
            return { status: 0, logId: Long.fromString('123'), prevId: Long.ZERO, sendAt: 1, msgId: 1 };
        }
    });
    const result = await session.sendChat(chat, true);
    assert.equal(sent.method, 'WRITE');
    assert.equal(sent.payload.chatId.toString(), roomId);
    assert.equal(sent.payload.msg, '@동명이인 확인');
    assert.equal(sent.payload.type, 1);
    assert.equal(sent.payload.noSeen, true);
    assert.ok(sent.payload.extra.includes('"user_id":' + firstId));
    assert.equal(result.result.logId.toString(), '123');
});

function fixture() {
    let sends = 0;
    const channel = {
        updateAll: async () => ({ success: true }),
        getUserInfo: ({ userId }) => userId.toString() === firstId ? { nickname: '현재이름', userId } : undefined,
        getDisplayName: () => '테스트 방',
        sendChat: async () => { sends++; return { success: true, result: { logId: Long.fromString('42') } }; }
    };
    const sender = new LocoSender('사용하지않음');
    sender.connected = true;
    sender.client = {
        logon: true, clientUser: { userId: Long.fromString('11') }, close() { this.logon = false; },
        channelList: { get: value => value.toString() === roomId ? channel : undefined }
    };
    return { sender, channel, count: () => sends };
}

test('방·대상 ID 확인, 최신 이름 적용, 준비만으로는 보내지 않음', async () => {
    const { sender, count } = fixture();
    await assert.rejects(sender.prepare({ chatId: '1', parts: [mention()] }), { code: 'CHANNEL_NOT_FOUND' });
    await assert.rejects(sender.prepare({ chatId: roomId, parts: [mention(secondId)] }), { code: 'USER_NOT_FOUND' });
    const prepared = await sender.prepare({ chatId: roomId, parts: [mention()] });
    assert.equal(prepared.text, '@현재이름');
    assert.equal(prepared.accountUserId, '11');
    assert.equal(count(), 0);
    await sender.sendPrepared(prepared.ticket);
    await assert.rejects(sender.sendPrepared(prepared.ticket), { code: 'INVALID_TICKET' });
    assert.equal(count(), 1);
});

test('만료·연결 종료·서버 거절·응답 손실에서 자동 재전송하지 않음', async () => {
    const { sender, channel } = fixture();
    let prepared = await sender.prepare({ chatId: roomId, parts: [mention()] });
    await assert.rejects(sender.prepare({ chatId: '1', parts: [mention()] }), { code: 'CHANNEL_NOT_FOUND' });
    await assert.rejects(sender.sendPrepared(prepared.ticket), { code: 'INVALID_TICKET' });
    prepared = await sender.prepare({ chatId: roomId, parts: [mention()] });
    sender.tickets.get(prepared.ticket).expires = 0;
    await assert.rejects(sender.sendPrepared(prepared.ticket), { code: 'INVALID_TICKET' });
    prepared = await sender.prepare({ chatId: roomId, parts: [mention()] });
    channel.sendChat = async () => { throw new Error('토큰이 포함될 수 있는 내부 오류'); };
    await assert.rejects(sender.sendPrepared(prepared.ticket), { code: 'SEND_UNKNOWN' });
    await assert.rejects(sender.sendPrepared(prepared.ticket), { code: 'INVALID_TICKET' });
    prepared = await sender.prepare({ chatId: roomId, parts: [mention()] });
    channel.sendChat = async () => ({ success: false, status: -1 });
    await assert.rejects(sender.sendPrepared(prepared.ticket), { code: 'WRITE_REJECTED', status: -1 });
    sender.close();
    await assert.rejects(sender.prepare({ chatId: roomId, parts: [mention()] }), { code: 'NOT_CONNECTED' });
});

test('강제 로그인 금지·기기 식별자 재사용·인증 정보 파일 미저장', async () => {
    const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'LocoSenderTest-'));
    const uuids = [];
    let force;
    let webConfig, locoConfig;
    const profile = { version: '26.8.1', appVersion: '26.8.1.5315' };
    class Client extends EventEmitter {
        constructor(config) { super(); locoConfig = config; this.logon = true; this.clientUser = { userId: Long.ONE }; this.channelList = { size: 1 }; }
        async login() { return { success: true }; }
        close() { this.logon = false; }
    }
    const sender = new LocoSender(directory, {
        AuthApiClient: { create: async (name, uuid, config) => { uuids.push(uuid); webConfig = config; return {
            login: async (form, forced) => { force = forced; return { success: true, result: { accessToken: '가짜토큰' } }; }
        }; } }, TalkClient: Client
    }, profile);
    try {
        await sender.login({ email: 'test@example.invalid', password: '가짜비밀번호' });
        assert.equal(force, false);
        assert.deepEqual(webConfig, profile);
        assert.deepEqual(locoConfig, profile);
        sender.close();
        await sender.auth();
        assert.equal(uuids[0], uuids[1]);
        assert.deepEqual(fs.readdirSync(directory), ['device.json']);
        assert.deepEqual(Object.keys(JSON.parse(fs.readFileSync(path.join(directory, 'device.json'), 'utf8'))), ['uuid']);
    } finally { sender.close(); fs.rmSync(directory, { recursive: true }); }
});

test('설치 버전 설정에서 인증 버전과 앱 버전을 일관되게 생성', () => {
    const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'LocoProfileTest-'));
    const file = path.join(directory, 'client-profile.json');
    try {
        fs.writeFileSync(file, JSON.stringify({ appVersion: '26.8.1.5315', locoBookingHost: '잘못된호스트' }));
        assert.deepEqual(loadProfile(file), { version: '26.8.1', appVersion: '26.8.1.5315' });
        fs.writeFileSync(file, JSON.stringify({ appVersion: '임의값' }));
        assert.throws(() => loadProfile(file), { code: 'CLIENT_PROFILE' });
    } finally { fs.rmSync(directory, { recursive: true }); }
});

test('미지원 등록 경로로 요청하지 않고 -100 이후 반복 로그인을 막음', async () => {
    let calls = 0;
    const sender = new LocoSender('사용하지않음');
    sender.auth = async () => {
        calls++;
        return { login: async () => ({ success: false, status: -100 }) };
    };
    assert.equal((await sender.handle({ op: 'info' })).registrationSupported, false);
    await assert.rejects(sender.handle({ op: 'requestPasscode' }), { code: 'REGISTRATION_UNSUPPORTED' });
    await assert.rejects(sender.handle({ op: 'registerDevice', passcode: '1234' }), { code: 'REGISTRATION_UNSUPPORTED' });
    assert.equal(calls, 0);
    const form = { email: 'test@example.invalid', password: '가짜비밀번호' };
    await assert.rejects(sender.login(form), { code: 'REGISTRATION_UNSUPPORTED', status: -100 });
    sender.close();
    await assert.rejects(sender.login(form), { code: 'REGISTRATION_UNSUPPORTED' });
    assert.equal(calls, 1);
});

test('실제 자식 프로세스의 UTF-8·요청 ID·오류 응답', async () => {
    const child = spawn(process.execPath, [path.join(__dirname, '..', 'bridge.js')], { stdio: ['pipe', 'pipe', 'pipe'] });
    try {
        const readline = require('node:readline').createInterface({ input: child.stdout });
        const iterator = readline[Symbol.asyncIterator]();
        child.stdin.write(JSON.stringify({ id: '한글요청', op: 'preview', chatId: roomId, parts: [mention(), text(' 테스트')] }) + '\n');
        const response = JSON.parse((await iterator.next()).value);
        assert.equal(response.id, '한글요청');
        assert.equal(response.result.text, '@동명이인 테스트');
        child.stdin.write(JSON.stringify({ id: '미로그인', op: 'prepare', chatId: roomId, parts: [mention()] }) + '\n');
        assert.equal(JSON.parse((await iterator.next()).value).error.code, 'NOT_CONNECTED');
        readline.close();
    } finally { child.kill(); }
});
