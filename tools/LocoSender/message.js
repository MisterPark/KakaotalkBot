'use strict';

const { Long, ChatBuilder, MentionContent, KnownChatType, util } = require('node-kakao');

class SenderError extends Error {
    constructor(code, message, status) {
        super(message);
        this.code = code;
        this.status = Number.isInteger(status) ? status : undefined;
    }
}

function id(value, field) {
    // JSON 경계에서는 숫자를 받지 않습니다. 2^53을 넘는 ID도 정확하게 유지합니다.
    if (typeof value !== 'string' || !/^[1-9]\d{0,18}$/.test(value) ||
        BigInt(value) > 9223372036854775807n) {
        throw new SenderError('INVALID_ID', `${field}: 양의 64비트 정수를 문자열로 전달해 주세요.`);
    }
    return Long.fromString(value);
}

function buildMessage(chatId, parts) {
    const channelId = id(chatId, 'chatId');
    if (!Array.isArray(parts) || parts.length === 0 || parts.length > 64) {
        throw new SenderError('INVALID_PARTS', '메시지 조각은 1~64개여야 합니다.');
    }
    const hasMention = parts.some(p => p && p.kind === 'mention');
    const builder = new ChatBuilder();
    for (const part of parts) {
        if (part && part.kind === 'text' && typeof part.text === 'string') {
            // 일반 @와 실제 멘션이 섞인 경우의 서버 해석은 아직 검증되지 않았습니다.
            if (hasMention && part.text.includes('@')) {
                throw new SenderError('AMBIGUOUS_AT', '멘션 메시지의 일반 본문에는 @를 넣을 수 없습니다.');
            }
            builder.text(part.text);
        } else if (part && part.kind === 'mention' && typeof part.nickname === 'string') {
            if (!part.nickname.trim() || /[@\r\n\x00]/.test(part.nickname) || part.nickname.length > 100) {
                throw new SenderError('INVALID_NICKNAME', '멘션 표시 이름이 비어 있거나 지원하지 않는 문자/길이입니다.');
            }
            builder.append(new MentionContent({ userId: id(part.userId, 'userId'), nickname: part.nickname }));
        } else {
            throw new SenderError('INVALID_PARTS', 'text 또는 mention 형식의 메시지 조각이 필요합니다.');
        }
    }
    const chat = builder.build(KnownChatType.TEXT);
    if (!chat.text.trim() || chat.text.length > 4000 || chat.text.includes('\0')) {
        throw new SenderError('INVALID_TEXT', '이 도구는 1~4000자의 본문만 지원합니다.');
    }
    return { channelId, chat };
}

function preview(chatId, parts) {
    const { chat } = buildMessage(chatId, parts);
    return {
        chatId, text: chat.text,
        mentions: (chat.attachment.mentions || []).map(m => ({
            userId: m.user_id.toString(), at: m.at, length: m.len
        })),
        // 실제 WRITE의 extra와 같은 직렬화입니다. ID를 JSON 숫자로 정확히 보냅니다.
        extra: util.JsonUtil.stringifyLoseless(chat.attachment)
    };
}

module.exports = { SenderError, id, buildMessage, preview };
