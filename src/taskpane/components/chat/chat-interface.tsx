import { MessageSquare, Moon, Settings, Sun, Trash2 } from "lucide-react";
import { type ReactNode, useState } from "react";
import { ChatProvider, useChat } from "./chat-context";
import { ChatInput } from "./chat-input";
import { MessageList } from "./message-list";
import { SettingsPanel } from "./settings-panel";
import type { ChatTab } from "./types";

type Theme = "light" | "dark";
const THEME_KEY = "openexcel-theme";

function useTheme() {
  const [theme, setTheme] = useState<Theme>(() => {
    const saved = localStorage.getItem(THEME_KEY) as Theme | null;
    const initial = saved ?? (window.matchMedia("(prefers-color-scheme: light)").matches ? "light" : "dark");
    document.documentElement.setAttribute("data-theme", initial);
    return initial;
  });

  const toggle = () => {
    const next = theme === "dark" ? "light" : "dark";
    document.documentElement.setAttribute("data-theme", next);
    localStorage.setItem(THEME_KEY, next);
    setTheme(next);
  };

  return { theme, toggle };
}

function formatTokens(n: number): string {
  if (n >= 1_000_000) return `${(n / 1_000_000).toFixed(1)}M`;
  if (n >= 1_000) return `${(n / 1_000).toFixed(1)}k`;
  return n.toString();
}

function formatCost(n: number): string {
  if (n < 0.01) return `$${n.toFixed(4)}`;
  return `$${n.toFixed(3)}`;
}

function StatsBar() {
  const { state } = useChat();
  const { sessionStats, providerConfig } = state;

  if (!providerConfig) return null;

  const totalTokens = sessionStats.inputTokens + sessionStats.outputTokens;
  const contextPct =
    sessionStats.contextWindow > 0 ? ((totalTokens / sessionStats.contextWindow) * 100).toFixed(1) : "0";

  return (
    <div
      className="flex items-center justify-between px-3 py-1.5 text-[10px] border-t border-(--chat-border) bg-(--chat-bg-secondary) text-(--chat-text-muted)"
      style={{ fontFamily: "var(--chat-font-mono)" }}
    >
      <div className="flex items-center gap-3">
        <span title="Input tokens">↑{formatTokens(sessionStats.inputTokens)}</span>
        <span title="Output tokens">↓{formatTokens(sessionStats.outputTokens)}</span>
        {sessionStats.cacheRead > 0 && <span title="Cache read tokens">R{formatTokens(sessionStats.cacheRead)}</span>}
        {sessionStats.cacheWrite > 0 && (
          <span title="Cache write tokens">W{formatTokens(sessionStats.cacheWrite)}</span>
        )}
        <span title="Total cost">{formatCost(sessionStats.totalCost)}</span>
        {sessionStats.contextWindow > 0 && (
          <span title="Context usage">
            {contextPct}%/{formatTokens(sessionStats.contextWindow)}
          </span>
        )}
      </div>
      <div className="flex items-center gap-1">
        <span>{providerConfig.provider}</span>
        <span className="text-(--chat-text-secondary)">{providerConfig.model}</span>
        {providerConfig.thinking !== "none" && (
          <span className="text-(--chat-accent)">• {providerConfig.thinking}</span>
        )}
      </div>
    </div>
  );
}

function TabButton({ active, onClick, children }: { active: boolean; onClick: () => void; children: ReactNode }) {
  return (
    <button
      type="button"
      onClick={onClick}
      className={`
        flex items-center gap-1.5 px-3 py-2 text-xs uppercase tracking-wider
        border-b-2 transition-colors
        ${
          active
            ? "border-(--chat-accent) text-(--chat-text-primary)"
            : "border-transparent text-(--chat-text-muted) hover:text-(--chat-text-secondary)"
        }
      `}
      style={{ fontFamily: "var(--chat-font-mono)" }}
    >
      {children}
    </button>
  );
}

function ChatHeader({
  activeTab,
  onTabChange,
  theme,
  onThemeToggle,
}: { activeTab: ChatTab; onTabChange: (tab: ChatTab) => void; theme: Theme; onThemeToggle: () => void }) {
  const { clearMessages, state } = useChat();

  return (
    <div className="border-b border-(--chat-border) bg-(--chat-bg)">
      <div className="flex items-center justify-between px-2">
        <div className="flex">
          <TabButton active={activeTab === "chat"} onClick={() => onTabChange("chat")}>
            <MessageSquare size={12} />
            Chat
          </TabButton>
          <TabButton active={activeTab === "settings"} onClick={() => onTabChange("settings")}>
            <Settings size={12} />
            Settings
          </TabButton>
        </div>
        <div className="flex items-center">
          <button
            type="button"
            onClick={onThemeToggle}
            className="p-1.5 text-(--chat-text-muted) hover:text-(--chat-text-primary) transition-colors"
            title={theme === "dark" ? "Switch to light mode" : "Switch to dark mode"}
          >
            {theme === "dark" ? <Sun size={14} /> : <Moon size={14} />}
          </button>
          {activeTab === "chat" && state.messages.length > 0 && (
            <button
              type="button"
              onClick={clearMessages}
              className="p-1.5 text-(--chat-text-muted) hover:text-(--chat-error) transition-colors"
              title="Clear messages"
            >
              <Trash2 size={14} />
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

function ChatContent() {
  const [activeTab, setActiveTab] = useState<ChatTab>("chat");
  const { theme, toggle } = useTheme();

  return (
    <div className="flex flex-col h-full bg-(--chat-bg)" style={{ fontFamily: "var(--chat-font-mono)" }}>
      <ChatHeader activeTab={activeTab} onTabChange={setActiveTab} theme={theme} onThemeToggle={toggle} />
      {activeTab === "chat" ? (
        <>
          <MessageList />
          <ChatInput />
          <StatsBar />
        </>
      ) : (
        <SettingsPanel />
      )}
    </div>
  );
}

export function ChatInterface() {
  return (
    <ChatProvider>
      <ChatContent />
    </ChatProvider>
  );
}
