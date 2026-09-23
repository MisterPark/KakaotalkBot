'use strict';

const path = require('node:path');
const { LocoSender } = require('./sender');
const { SenderError } = require('./message');

// 표준 입출력만 사용합니다. HTTP 포트를 열거나 입력 요청/인증 응답을 기록하지 않습니다.
const sender = new LocoSender(process.argv[2] || path.join(__dirname, 'data'));
let buffer = '';
let busy = false;
function write(value) { process.stdout.write(JSON.stringify(value) + '\n'); }
function terminate() { sender.close(); process.exit(0); }
process.stdin.setEncoding('utf8');
process.stdin.on('end', terminate);
process.on('SIGTERM', terminate);
process.on('uncaughtException', () => { process.stderr.write('LOCO 보조 프로세스 오류\n'); process.exit(1); });
process.on('unhandledRejection', () => { process.stderr.write('LOCO 비동기 처리 오류\n'); process.exit(1); });

async function handle(line) {
    let request;
    try { request = JSON.parse(line); } catch (_) { write({ id: null, ok: false, error: { code: 'JSON', message: '올바른 JSON 요청이 아닙니다.' } }); return; }
    if (!request || typeof request.id !== 'string' || request.id.length > 100) {
        write({ id: null, ok: false, error: { code: 'ID', message: '요청 식별자가 필요합니다.' } }); return;
    }
    if (busy) { write({ id: request.id, ok: false, error: { code: 'BUSY', message: '이전 요청 처리 중입니다.' } }); return; }
    busy = true;
    // 시간 제한 이후에는 소켓과 프로세스를 종료합니다. 늦게 도착한 응답은 재전송의 근거가 아닙니다.
    const timer = setTimeout(() => {
        write({ id: request.id, ok: false, error: { code: request.op === 'sendPrepared' ? 'SEND_UNKNOWN' : 'TIMEOUT', message: '요청 시간이 초과되었습니다. 전송 요청이었다면 채팅방에서 결과를 확인해 주세요.' } });
        process.stdout.write('', () => { sender.close(); process.exit(2); });
    }, 30000);
    try {
        write({ id: request.id, ok: true, result: await sender.handle(request) });
    } catch (error) {
        // 외부 라이브러리 오류에는 비밀번호/토큰/HTTP 본문이 포함될 수 있습니다.
        const safe = error instanceof SenderError ? { code: error.code, message: error.message, status: error.status } :
            { code: 'TRANSPORT', message: '서버 통신 또는 라이브러리 처리에 실패했습니다.', httpStatus: Number.isInteger(error?.response?.status) ? error.response.status : undefined };
        write({ id: request.id, ok: false, error: safe });
    } finally {
        request.password = undefined;
        request.passcode = undefined;
        clearTimeout(timer);
        busy = false;
    }
}

process.stdin.on('data', chunk => {
    buffer += chunk;
    if (Buffer.byteLength(buffer, 'utf8') > 65536) { process.stderr.write('입력 크기 제한 초과\n'); terminate(); return; }
    let newline;
    while ((newline = buffer.indexOf('\n')) >= 0) {
        const line = buffer.slice(0, newline);
        buffer = buffer.slice(newline + 1);
        if (line.trim()) void handle(line);
    }
});
