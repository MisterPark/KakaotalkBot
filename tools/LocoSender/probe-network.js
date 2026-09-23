'use strict';

const net = require('node:net');
const https = require('node:https');

// 공개 서버의 연결 가능 여부만 검사합니다. 인증 요청이나 메시지 전송은 하지 않습니다.
function tcp(host) {
    return new Promise(resolve => {
        const socket = net.connect({ host, port: 443 });
        const finish = result => { socket.destroy(); resolve({ host, ...result }); };
        socket.setTimeout(8000, () => finish({ tcp: false, code: 'TIMEOUT' }));
        socket.once('connect', () => finish({ tcp: true }));
        socket.once('error', error => finish({ tcp: false, code: error.code }));
    });
}
function web() {
    return new Promise(resolve => {
        const request = https.request('https://katalk.kakao.com/', { method: 'HEAD', timeout: 8000 }, response => {
            response.resume();
            resolve({ host: 'katalk.kakao.com', tls: true, httpStatus: response.statusCode });
        });
        request.on('timeout', () => request.destroy(new Error('TIMEOUT')));
        request.on('error', error => resolve({ host: 'katalk.kakao.com', tls: false, code: error.code || 'TIMEOUT' }));
        request.end();
    });
}
Promise.all([tcp('booking-loco.kakao.com'), web()]).then(results => {
    console.log(JSON.stringify({ checkedAt: new Date().toISOString(), results, authenticated: false, messageSent: false }, null, 2));
});
