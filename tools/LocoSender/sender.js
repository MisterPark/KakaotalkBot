'use strict';

const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const kakao = require('node-kakao');
const { SenderError, id, buildMessage, preview } = require('./message');
const { loadProfile } = require('./profile');

// 현재 서버에서 기존 인증번호 요청 경로가 HTTP 404를 반환했습니다.
// 대체 등록 절차를 검증하기 전에는 같은 요청이나 임의 경로로 재시도하지 않습니다.
const registrationUnavailable = '기존 인증번호 요청 API에서 HTTP 404가 확인됐습니다. 현재 LOCO 구현은 신규 기기 등록을 지원하지 않습니다.';

function requireSuccess(result, stage) {
    if (result && result.success) return result.result;
    const status = result && result.status;
    const details = {
        '-100': '서버가 이 기기를 미등록으로 판단했습니다. 현재 LOCO 구현의 신규 기기 등록 경로는 사용할 수 없습니다.',
        '-101': '다른 기기가 로그인 중입니다. 강제 로그인은 하지 않았습니다.',
        '-999': '클라이언트 버전 갱신이 필요합니다. 현재 라이브러리와 서버의 호환성을 확인해야 합니다.',
        '12': '계정 또는 비밀번호를 확인해 주세요.',
        '13': '로그인 시도 제한입니다. 재시도를 멈추고 카카오톡에서 확인해 주세요.',
        '-111': '인증번호가 올바르지 않습니다.'
    };
    throw new SenderError(stage, details[String(status)] || `${stage}: 서버가 요청을 거절했습니다.`, status);
}

class LocoSender {
    constructor(dataDirectory, dependencies = kakao, config = loadProfile()) {
        this.dependencies = dependencies;
        this.config = { ...config };
        this.dataDirectory = dataDirectory;
        this.client = null;
        this.connected = false;
        this.tickets = new Map();
        this.unregistered = false;
    }

    async auth() {
        // 비밀번호와 액세스 토큰은 저장하지 않고, 이 보조 기기의 식별자만 보존합니다.
        fs.mkdirSync(this.dataDirectory, { recursive: true });
        const file = path.join(this.dataDirectory, 'device.json');
        if (!fs.existsSync(file)) {
            fs.writeFileSync(file, JSON.stringify({ uuid: crypto.randomBytes(64).toString('base64') }), { flag: 'wx' });
        }
        const device = JSON.parse(fs.readFileSync(file, 'utf8'));
        if (typeof device.uuid !== 'string' || !/^[A-Za-z0-9+/]{86}==$/.test(device.uuid)) {
            throw new SenderError('DEVICE_FILE', '기기 식별자 파일이 손상되었습니다.');
        }
        return this.dependencies.AuthApiClient.create('KakaotalkBot LOCO 진단', device.uuid, this.config);
    }

    credentials(request) {
        if (typeof request.email !== 'string' || !request.email.trim() ||
            typeof request.password !== 'string' || !request.password) {
            throw new SenderError('CREDENTIALS', '계정과 비밀번호를 진단 창에 입력해 주세요.');
        }
        return { email: request.email.trim(), password: request.password };
    }

    close() {
        const client = this.client;
        this.client = null;
        this.connected = false;
        this.tickets.clear();
        if (client && client.logon) {
            try { client.close(); } catch (_) { /* 종료 중인 소켓은 다시 닫지 않습니다. */ }
        }
    }

    async login(request) {
        if (this.unregistered) throw new SenderError('REGISTRATION_UNSUPPORTED', registrationUnavailable);
        if (this.connected) throw new SenderError('ALREADY_CONNECTED', '이미 연결되어 있습니다.');
        this.close();
        const api = await this.auth();
        const loginResult = await api.login(this.credentials(request), false);
        if (loginResult && !loginResult.success && loginResult.status === -100) {
            this.unregistered = true;
            throw new SenderError('REGISTRATION_UNSUPPORTED', registrationUnavailable, -100);
        }
        const credential = requireSuccess(loginResult, 'WEB_LOGIN');
        const client = this.client = new this.dependencies.TalkClient(this.config);
        // 연결이 끊어지면 준비한 전송도 폐기합니다. 자동 재로그인/재전송은 하지 않습니다.
        const disconnected = () => { this.connected = false; this.tickets.clear(); };
        client.on('error', disconnected);
        client.on('disconnected', disconnected);
        client.on('switch_server', () => this.close());
        try {
            requireSuccess(await client.login(credential), 'LOCO_LOGIN');
            this.connected = true;
            return { userId: client.clientUser.userId.toString(), channelCount: client.channelList.size };
        } catch (error) {
            this.close();
            throw error;
        }
    }

    async prepare(request) {
        this.tickets.clear();
        const original = buildMessage(request.chatId, request.parts);
        if (!this.connected || !this.client || !this.client.logon) {
            throw new SenderError('NOT_CONNECTED', 'LOCO 로그인이 필요합니다.');
        }
        const channel = this.client.channelList.get(original.channelId);
        if (!channel) throw new SenderError('CHANNEL_NOT_FOUND', '로그인 계정에서 해당 채팅방 ID를 찾지 못했습니다.');
        requireSuccess(await channel.updateAll(), 'CHANNEL_UPDATE');
        const parts = request.parts.map(part => {
            if (part.kind !== 'mention') return { kind: 'text', text: part.text };
            const user = channel.getUserInfo({ userId: id(part.userId, 'userId') });
            if (!user) throw new SenderError('USER_NOT_FOUND', '해당 채팅방에서 멘션 대상 ID를 확인하지 못했습니다.');
            // 표시 이름도 해당 방의 최신 사용자 정보에서 얻습니다. 동명이인은 ID로 구별합니다.
            return { kind: 'mention', userId: part.userId, nickname: user.nickname };
        });
        const result = preview(request.chatId, parts);
        const ticket = crypto.randomUUID();
        this.tickets.set(ticket, { channel, parts, chatId: request.chatId, expires: Date.now() + 120000 });
        return { ...result, ticket, accountUserId: this.client.clientUser.userId.toString(), channelName: channel.getDisplayName() };
    }

    async sendPrepared(ticket) {
        const pending = this.tickets.get(ticket);
        // 서버 응답을 잃어도 같은 티켓으로 다시 보내지 않도록 전송 전에 소비합니다.
        this.tickets.delete(ticket);
        if (!pending || pending.expires < Date.now()) throw new SenderError('INVALID_TICKET', '전송 준비가 만료되었습니다. 다시 미리보기를 확인해 주세요.');
        if (!this.connected || !this.client || !this.client.logon) throw new SenderError('NOT_CONNECTED', 'LOCO 연결이 끊겼습니다.');
        const { chat } = buildMessage(pending.chatId, pending.parts);
        let result;
        try {
            result = await pending.channel.sendChat(chat, true);
        } catch (_) {
            throw new SenderError('SEND_UNKNOWN', '전송 결과를 확인하지 못했습니다. 채팅방을 확인하기 전에는 다시 보내지 마세요.');
        }
        const log = requireSuccess(result, 'WRITE_REJECTED');
        return { chatId: pending.chatId, logId: log && log.logId ? log.logId.toString() : null, accepted: true };
    }

    async handle(request) {
        switch (request.op) {
            case 'info': return { protocol: 1, library: 'node-kakao 4.5.0', appVersion: this.config.appVersion, version: this.config.version, registrationSupported: false, registrationNotice: registrationUnavailable, connected: this.connected && !!this.client && this.client.logon };
            case 'preview': return preview(request.chatId, request.parts);
            case 'login': return this.login(request);
            case 'requestPasscode':
            case 'registerDevice': throw new SenderError('REGISTRATION_UNSUPPORTED', registrationUnavailable);
            case 'prepare': return this.prepare(request);
            case 'sendPrepared': return this.sendPrepared(request.ticket);
            case 'close': this.close(); return { closed: true };
            default: throw new SenderError('UNKNOWN_OPERATION', '지원하지 않는 요청입니다.');
        }
    }
}

module.exports = { LocoSender };
