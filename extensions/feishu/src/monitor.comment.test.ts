import { hasControlCommand } from "openclaw/plugin-sdk/command-auth";
import {
  createInboundDebouncer,
  resolveInboundDebounceMs,
} from "openclaw/plugin-sdk/reply-runtime";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createPluginRuntimeMock } from "../../../test/helpers/plugins/plugin-runtime-mock.js";
import { createNonExitingTypedRuntimeEnv } from "../../../test/helpers/plugins/runtime-env.js";
import type { ClawdbotConfig, RuntimeEnv } from "../runtime-api.js";
import type { FeishuMessageEvent } from "./bot.js";
import { monitorSingleAccount } from "./monitor.account.js";
import {
  resolveDriveCommentSyntheticEvent,
  type FeishuDriveCommentNoticeEvent,
} from "./monitor.comment.js";
import { setFeishuRuntime } from "./runtime.js";
import type { ResolvedFeishuAccount } from "./types.js";

const handleFeishuMessageMock = vi.hoisted(() => vi.fn(async (_params: { event?: unknown }) => {}));
const createEventDispatcherMock = vi.hoisted(() => vi.fn());
const createFeishuClientMock = vi.hoisted(() => vi.fn());
const monitorWebSocketMock = vi.hoisted(() => vi.fn(async () => {}));
const monitorWebhookMock = vi.hoisted(() => vi.fn(async () => {}));
const createFeishuThreadBindingManagerMock = vi.hoisted(() => vi.fn(() => ({ stop: vi.fn() })));

let handlers: Record<string, (data: unknown) => Promise<void>> = {};

vi.mock("./client.js", () => ({
  createEventDispatcher: createEventDispatcherMock,
  createFeishuClient: createFeishuClientMock,
}));

vi.mock("./bot.js", async () => {
  const actual = await vi.importActual<typeof import("./bot.js")>("./bot.js");
  return {
    ...actual,
    handleFeishuMessage: handleFeishuMessageMock,
  };
});

vi.mock("./monitor.transport.js", () => ({
  monitorWebSocket: monitorWebSocketMock,
  monitorWebhook: monitorWebhookMock,
}));

vi.mock("./thread-bindings.js", () => ({
  createFeishuThreadBindingManager: createFeishuThreadBindingManagerMock,
}));

function buildMonitorConfig(): ClawdbotConfig {
  return {
    channels: {
      feishu: {
        enabled: true,
      },
    },
  } as ClawdbotConfig;
}

function buildMonitorAccount(): ResolvedFeishuAccount {
  return {
    accountId: "default",
    enabled: true,
    configured: true,
    appId: "cli_test",
    appSecret: "secret_test", // pragma: allowlist secret
    domain: "feishu",
    config: {
      enabled: true,
      connectionMode: "websocket",
    },
  } as ResolvedFeishuAccount;
}

function makeDriveCommentEvent(
  overrides: Partial<FeishuDriveCommentNoticeEvent> = {},
): FeishuDriveCommentNoticeEvent {
  return {
    comment_id: "7623358762119646411",
    event_id: "10d9d60b990db39f96a4c2fd357fb877",
    is_mentioned: true,
    notice_meta: {
      file_token: "GS9sdtIlOonqKSx2PtrcRqgCnBe",
      file_type: "docx",
      from_user_id: {
        open_id: "ou_509d4d7ace4a9addec2312676ffcba9b",
      },
      notice_type: "add_comment",
      to_user_id: {
        open_id: "ou_bot",
      },
    },
    reply_id: "7623358762136374451",
    timestamp: "1774951528000",
    type: "drive.notice.comment_add_v1",
    ...overrides,
  };
}

function makeOpenApiClient(params: {
  documentTitle?: string;
  documentUrl?: string;
  quoteText?: string;
  rootReplyText?: string;
  targetReplyText?: string;
  includeTargetReplyInBatch?: boolean;
}) {
  return {
    request: vi.fn(async (request: { method: "GET" | "POST"; url: string; data: unknown }) => {
      if (request.url === "/open-apis/drive/v1/metas/batch_query") {
        return {
          code: 0,
          data: {
            metas: [
              {
                doc_token: "GS9sdtIlOonqKSx2PtrcRqgCnBe",
                title: params.documentTitle ?? "评论事件处理需求",
                url:
                  params.documentUrl ??
                  "https://bytedance.larkoffice.com/docx/ZE56dawdDoGb0xxxVcbcmpfinZc",
              },
            ],
          },
        };
      }
      if (request.url.includes("/comments/batch_query")) {
        return {
          code: 0,
          data: {
            items: [
              {
                comment_id: "7623358762119646411",
                quote: params.quoteText ?? "im.message.receive_v1 消息触发实现",
                reply_list: {
                  replies: [
                    {
                      reply_id: "7623358762136374451",
                      content: {
                        elements: [
                          {
                            type: "text_run",
                            text_run: {
                              content: params.rootReplyText ?? "收到评论事件后，也发送给agent",
                            },
                          },
                        ],
                      },
                    },
                    ...(params.includeTargetReplyInBatch
                      ? [
                          {
                            reply_id: "7623359125036043462",
                            content: {
                              elements: [
                                {
                                  type: "text_run",
                                  text_run: {
                                    content: params.targetReplyText ?? "跟进处理一下这个评论",
                                  },
                                },
                              ],
                            },
                          },
                        ]
                      : []),
                  ],
                },
              },
            ],
          },
        };
      }
      if (request.url.includes("/replies")) {
        return {
          code: 0,
          data: {
            has_more: false,
            items: [
              {
                reply_id: "7623358762136374451",
                content: {
                  elements: [
                    {
                      type: "text_run",
                      text_run: {
                        content: params.rootReplyText ?? "收到评论事件后，也发送给agent",
                      },
                    },
                  ],
                },
              },
              {
                reply_id: "7623359125036043462",
                content: {
                  elements: [
                    {
                      type: "text_run",
                      text_run: {
                        content: params.targetReplyText ?? "跟进处理一下这个评论",
                      },
                    },
                  ],
                },
              },
            ],
          },
        };
      }
      throw new Error(`unexpected request: ${request.method} ${request.url}`);
    }),
  };
}

async function setupCommentMonitorHandler(): Promise<(data: unknown) => Promise<void>> {
  const register = vi.fn((registered: Record<string, (data: unknown) => Promise<void>>) => {
    handlers = registered;
  });
  createEventDispatcherMock.mockReturnValue({ register });

  await monitorSingleAccount({
    cfg: buildMonitorConfig(),
    account: buildMonitorAccount(),
    runtime: createNonExitingTypedRuntimeEnv<RuntimeEnv>(),
    botOpenIdSource: {
      kind: "prefetched",
      botOpenId: "ou_bot",
    },
  });

  const handler = handlers["drive.notice.comment_add_v1"];
  if (!handler) {
    throw new Error("missing drive.notice.comment_add_v1 handler");
  }
  return handler;
}

function extractSyntheticText(event: FeishuMessageEvent): string {
  const content = JSON.parse(event.message.content) as { text?: string };
  return content.text ?? "";
}

describe("resolveDriveCommentSyntheticEvent", () => {
  it("builds a synthetic Feishu message for add_comment notices", async () => {
    const client = makeOpenApiClient({ includeTargetReplyInBatch: true });

    const synthetic = await resolveDriveCommentSyntheticEvent({
      cfg: buildMonitorConfig(),
      accountId: "default",
      event: makeDriveCommentEvent(),
      botOpenId: "ou_bot",
      createClient: () => client as never,
    });

    expect(synthetic).not.toBeNull();
    expect(synthetic?.sender.sender_id.open_id).toBe("ou_509d4d7ace4a9addec2312676ffcba9b");
    expect(synthetic?.message.message_id).toBe("drive-comment:10d9d60b990db39f96a4c2fd357fb877");
    expect(synthetic?.message.chat_id).toBe("p2p:ou_509d4d7ace4a9addec2312676ffcba9b");
    expect(synthetic?.message.chat_type).toBe("p2p");
    expect(synthetic?.message.create_time).toBe("1774951528000");

    const text = extractSyntheticText(synthetic as FeishuMessageEvent);
    expect(text).toContain("《评论事件处理需求》");
    expect(text).toContain("添加了一条评论");
    expect(text).toContain("收到评论事件后，也发送给agent");
    expect(text).toContain("评论引用内容：im.message.receive_v1 消息触发实现");
    expect(text).toContain("这条评论提到了你。");
    expect(text).toContain("这是飞书文档评论事件，不是普通即时消息对话。");
    expect(text).toContain("feishu_drive.reply_comment");
    expect(text).toContain("最终请通过 feishu_drive.reply_comment 在该评论线程中回复答案");
    expect(text).toContain(
      "也请在完成后通过 feishu_drive.reply_comment 在该评论线程中告知用户已修改完成",
    );
    expect(text).toContain("最终请只输出 NO_REPLY");
  });

  it("falls back to the replies API to resolve add_reply text", async () => {
    const client = makeOpenApiClient({
      includeTargetReplyInBatch: false,
      targetReplyText: "跟进处理一下这个评论",
    });

    const synthetic = await resolveDriveCommentSyntheticEvent({
      cfg: buildMonitorConfig(),
      accountId: "default",
      event: makeDriveCommentEvent({
        notice_meta: {
          ...makeDriveCommentEvent().notice_meta,
          notice_type: "add_reply",
        },
        reply_id: "7623359125036043462",
      }),
      botOpenId: "ou_bot",
      createClient: () => client as never,
    });

    const text = extractSyntheticText(synthetic as FeishuMessageEvent);
    expect(text).toContain("添加了一条回复：跟进处理一下这个评论");
    expect(text).toContain("原评论：收到评论事件后，也发送给agent");
    expect(text).toContain("file_token：GS9sdtIlOonqKSx2PtrcRqgCnBe");
    expect(text).toContain("事件类型：add_reply");
  });

  it("ignores self-authored comment notices", async () => {
    const synthetic = await resolveDriveCommentSyntheticEvent({
      cfg: buildMonitorConfig(),
      accountId: "default",
      event: makeDriveCommentEvent({
        notice_meta: {
          ...makeDriveCommentEvent().notice_meta,
          from_user_id: { open_id: "ou_bot" },
        },
      }),
      botOpenId: "ou_bot",
      createClient: () => makeOpenApiClient({}) as never,
    });

    expect(synthetic).toBeNull();
  });
});

describe("drive.notice.comment_add_v1 monitor handler", () => {
  beforeEach(() => {
    handlers = {};
    handleFeishuMessageMock.mockClear();
    createEventDispatcherMock.mockReset();
    createFeishuClientMock.mockReset().mockReturnValue(makeOpenApiClient({}) as never);
    createFeishuThreadBindingManagerMock.mockReset().mockImplementation(() => ({
      stop: vi.fn(),
    }));
    setFeishuRuntime(
      createPluginRuntimeMock({
        channel: {
          debounce: {
            resolveInboundDebounceMs,
            createInboundDebouncer,
          },
          text: {
            hasControlCommand,
          },
        },
      }),
    );
  });

  afterEach(() => {
    vi.restoreAllMocks();
  });

  it("dispatches comment notices through handleFeishuMessage", async () => {
    const onComment = await setupCommentMonitorHandler();

    await onComment(makeDriveCommentEvent());

    expect(handleFeishuMessageMock).toHaveBeenCalledTimes(1);
    const params = handleFeishuMessageMock.mock.calls[0]?.[0] as
      | { event?: FeishuMessageEvent }
      | undefined;
    expect(params?.event).toBeDefined();
    const text = extractSyntheticText(params?.event as FeishuMessageEvent);
    expect(text).toContain("请根据这次文档评论事件决定接下来要做什么。");
  });
});
