/**
 * constants.js - 全局命名常量
 */

// 超时（毫秒）
export const TIMEOUT = {
  OFFICE_INIT: 3000,
  MODEL_FETCH: 15000,
  LLM_CALL: 60000,
  LLM_STREAM_IDLE: 30000,
  RETRY_DELAY: 500,
};

// 重试
export const RETRY = {
  LIMIT: 3,
};

// UI 延迟（毫秒）
export const UI_DELAY = {
  ACTION_DEBOUNCE: 1000,
  URL_ACTION: 800,
  LOADING_ANIMATION: 600,
  HIDE_STATUS: 200,
  TOAST_DISPLAY: 2500,
  TOAST_HIDE: 300,
  INLINE_STATUS_TIMEOUT: 5000,
};

// 并发限制
export const CONCURRENCY = {
  MIN: 1,
  MAX: 5,
  DEFAULT: 3,
};

// 模型防抖
export const MODEL_FETCH_DEBOUNCE_MS = 30000;