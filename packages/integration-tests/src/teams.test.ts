import { createMemoryState } from "@chat-adapter/state-memory";
import { createTeamsAdapter, type TeamsAdapter } from "@chat-adapter/teams";
import { Chat, type Logger } from "chat";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import {
  createMockTeamsApp,
  createTeamsActivity,
  createTeamsWebhookRequest,
  DEFAULT_TEAMS_SERVICE_URL,
  getTeamsThreadId,
  injectMockTeamsApp,
  type MockTeamsApp,
  TEAMS_APP_ID,
  TEAMS_APP_PASSWORD,
  TEAMS_BOT_ID,
  TEAMS_BOT_NAME,
} from "./teams-utils";
import { createWaitUntilTracker } from "./test-scenarios";

const HELP_REGEX = /help/i;
const ANY_CHAR_REGEX = /./;

const mockLogger: Logger = {
  debug: vi.fn(),
  info: vi.fn(),
  warn: vi.fn(),
  error: vi.fn(),
  child: () => mockLogger,
};

describe("Teams Integration", () => {
  let chat: Chat<{ teams: TeamsAdapter }>;
  let state: ReturnType<typeof createMemoryState>;
  let teamsAdapter: TeamsAdapter;
  let mockTeamsApp: MockTeamsApp;
  let tracker: ReturnType<typeof createWaitUntilTracker>;

  const TEST_CONVERSATION_ID = "19:meeting_123@thread.v2";
  const TEST_THREAD_ID = getTeamsThreadId(
    TEST_CONVERSATION_ID,
    DEFAULT_TEAMS_SERVICE_URL
  );

  beforeEach(() => {
    vi.clearAllMocks();

    state = createMemoryState();
    teamsAdapter = createTeamsAdapter({
      appId: TEAMS_APP_ID,
      appPassword: TEAMS_APP_PASSWORD,
      userName: TEAMS_BOT_NAME,
      logger: mockLogger,
    });

    mockTeamsApp = createMockTeamsApp();
    injectMockTeamsApp(teamsAdapter, mockTeamsApp);

    chat = new Chat({
      userName: TEAMS_BOT_NAME,
      adapters: { teams: teamsAdapter },
      state,
      logger: "error",
    });

    tracker = createWaitUntilTracker();
  });

  afterEach(async () => {
    await chat.shutdown();
  });

  describe("message handling", () => {
    it("should handle an @mention and call the handler", async () => {
      const handlerMock = vi.fn();
      chat.onNewMention(async (thread, message) => {
        handlerMock(thread.id, message.text);
        await thread.post("Hello from Teams!");
      });

      const activity = createTeamsActivity({
        text: `<at>${TEAMS_BOT_NAME}</at> hello bot!`,
        messageId: "msg-001",
        conversationId: TEST_CONVERSATION_ID,
        fromId: "user-123",
        fromName: "John Doe",
        mentions: [
          {
            id: TEAMS_BOT_ID,
            name: TEAMS_BOT_NAME,
            text: `<at>${TEAMS_BOT_NAME}</at>`,
          },
        ],
      });

      const request = createTeamsWebhookRequest(activity);
      const response = await chat.webhooks.teams(request, {
        waitUntil: tracker.waitUntil,
      });
      expect(response.status).toBe(200);

      await tracker.waitForAll();

      expect(handlerMock).toHaveBeenCalledWith(
        TEST_THREAD_ID,
        `@${TEAMS_BOT_NAME} hello bot!`
      );

      expect(mockTeamsApp.sentActivities.length).toBeGreaterThan(0);
    });

    it("should handle messages in subscribed threads", async () => {
      chat.onNewMention(async (thread) => {
        await thread.subscribe();
        await thread.post("I'm now listening!");
      });

      const subscribedHandler = vi.fn();
      chat.onSubscribedMessage(async (thread, message) => {
        subscribedHandler(thread.id, message.text);
        await thread.post(`You said: ${message.text}`);
      });

      // Initial mention to subscribe
      const mentionActivity = createTeamsActivity({
        text: `<at>${TEAMS_BOT_NAME}</at> subscribe me`,
        messageId: "msg-001",
        conversationId: TEST_CONVERSATION_ID,
        fromId: "user-123",
        fromName: "John Doe",
        mentions: [
          {
            id: TEAMS_BOT_ID,
            name: TEAMS_BOT_NAME,
            text: `<at>${TEAMS_BOT_NAME}</at>`,
          },
        ],
      });

      await chat.webhooks.teams(createTeamsWebhookRequest(mentionActivity), {
        waitUntil: tracker.waitUntil,
      });
      await tracker.waitForAll();

      mockTeamsApp.clearMocks();

      // Follow-up message in same thread
      const followUpActivity = createTeamsActivity({
        text: "This is a follow-up message",
        messageId: "msg-002",
        conversationId: TEST_CONVERSATION_ID,
        fromId: "user-123",
        fromName: "John Doe",
      });

      await chat.webhooks.teams(createTeamsWebhookRequest(followUpActivity), {
        waitUntil: tracker.waitUntil,
      });
      await tracker.waitForAll();

      expect(subscribedHandler).toHaveBeenCalledWith(
        TEST_THREAD_ID,
        "This is a follow-up message"
      );
    });

    it("should handle messages matching a pattern", async () => {
      const patternHandler = vi.fn();
      chat.onNewMessage(HELP_REGEX, async (thread, message) => {
        patternHandler(message.text);
        await thread.post("Here is some help!");
      });

      const activity = createTeamsActivity({
        text: "I need help with something",
        messageId: "msg-001",
        conversationId: TEST_CONVERSATION_ID,
        fromId: "user-123",
        fromName: "John Doe",
      });

      await chat.webhooks.teams(createTeamsWebhookRequest(activity), {
        waitUntil: tracker.waitUntil,
      });
      await tracker.waitForAll();

      expect(patternHandler).toHaveBeenCalledWith("I need help with something");
    });

    it("should skip messages from the bot itself", async () => {
      const handlerMock = vi.fn();
      chat.onNewMessage(ANY_CHAR_REGEX, () => {
        handlerMock();
      });

      const activity = createTeamsActivity({
        text: "Bot's own message",
        messageId: "msg-001",
        conversationId: TEST_CONVERSATION_ID,
        fromId: TEAMS_BOT_ID,
        fromName: TEAMS_BOT_NAME,
        isFromBot: true,
        recipientId: TEAMS_BOT_ID,
      });

      await chat.webhooks.teams(createTeamsWebhookRequest(activity), {
        waitUntil: tracker.waitUntil,
      });
      await tracker.waitForAll();

      expect(handlerMock).not.toHaveBeenCalled();
    });

    it("should skip non-message activity types", async () => {
      const handlerMock = vi.fn();
      chat.onNewMessage(ANY_CHAR_REGEX, () => {
        handlerMock();
      });

      const activity = createTeamsActivity({
        type: "conversationUpdate",
        text: "",
        messageId: "msg-001",
        conversationId: TEST_CONVERSATION_ID,
        fromId: "user-123",
        fromName: "John Doe",
      });

      await chat.webhooks.teams(createTeamsWebhookRequest(activity), {
        waitUntil: tracker.waitUntil,
      });
      await tracker.waitForAll();

      expect(handlerMock).not.toHaveBeenCalled();
    });
  });

  describe("thread operations", () => {
    it("should include thread info in message objects", async () => {
      let capturedMessage: unknown;
      chat.onNewMention((_thread, message) => {
        capturedMessage = message;
      });

      const activity = createTeamsActivity({
        text: `<at>${TEAMS_BOT_NAME}</at> test`,
        messageId: "msg-001",
        conversationId: TEST_CONVERSATION_ID,
        fromId: "user-123",
        fromName: "John Doe",
        mentions: [
          {
            id: TEAMS_BOT_ID,
            name: TEAMS_BOT_NAME,
            text: `<at>${TEAMS_BOT_NAME}</at>`,
          },
        ],
      });

      await chat.webhooks.teams(createTeamsWebhookRequest(activity), {
        waitUntil: tracker.waitUntil,
      });
      await tracker.waitForAll();

      expect(capturedMessage).toBeDefined();
      const msg = capturedMessage as {
        threadId: string;
        author: { userId: string; userName: string; isBot: boolean };
      };
      expect(msg.threadId).toBe(TEST_THREAD_ID);
      expect(msg.author.userId).toBe("user-123");
      expect(msg.author.userName).toBe("John Doe");
      expect(msg.author.isBot).toBe(false);
    });
  });

  describe("markdown formatting", () => {
    it("should convert markdown to Teams format in posted messages", async () => {
      chat.onNewMention(async (thread) => {
        await thread.post({
          markdown: "**Bold** and _italic_ and `code`",
        });
      });

      const activity = createTeamsActivity({
        text: `<at>${TEAMS_BOT_NAME}</at> markdown test`,
        messageId: "msg-001",
        conversationId: TEST_CONVERSATION_ID,
        fromId: "user-123",
        fromName: "John Doe",
        mentions: [
          {
            id: TEAMS_BOT_ID,
            name: TEAMS_BOT_NAME,
            text: `<at>${TEAMS_BOT_NAME}</at>`,
          },
        ],
      });

      await chat.webhooks.teams(createTeamsWebhookRequest(activity), {
        waitUntil: tracker.waitUntil,
      });
      await tracker.waitForAll();

      // Check that sent activities contain markdown
      const sentWithText = mockTeamsApp.sentActivities.find(
        (act: unknown) =>
          typeof act === "object" && act !== null && "text" in act
      );
      expect(sentWithText).toBeDefined();
      const text = (sentWithText as { text: string }).text;
      expect(text).toContain("**Bold**");
      expect(text).toContain("_italic_");
      expect(text).toContain("`code`");
    });
  });
});
