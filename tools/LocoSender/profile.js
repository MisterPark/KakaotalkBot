'use strict';

const fs = require('node:fs');
const path = require('node:path');
const { DefaultConfiguration } = require('node-kakao');
const { SenderError } = require('./message');

function loadProfile(file = path.join(__dirname, 'client-profile.json')) {
    if (!fs.existsSync(file)) return { version: DefaultConfiguration.version, appVersion: DefaultConfiguration.appVersion };
    const value = JSON.parse(fs.readFileSync(file, 'utf8'));
    if (typeof value.appVersion !== 'string' || !/^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,6}$/.test(value.appVersion)) {
        throw new SenderError('CLIENT_PROFILE', '설치된 카카오톡 버전 설정을 읽지 못했습니다. 버전 확인 스크립트를 실행해 주세요.');
    }
    // 인증 서버와 LOCO 접속에 같은 버전을 전달합니다. 서버 주소는 설정으로 변경하지 않습니다.
    return { version: value.appVersion.split('.').slice(0, 3).join('.'), appVersion: value.appVersion };
}

module.exports = { loadProfile };
