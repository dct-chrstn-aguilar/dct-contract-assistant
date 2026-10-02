import type { ReactNode } from "react";
import AssetIcon, { type AssetIconName } from "./AssetIcon";

export type Section = "chat" | "jira" | "sales";

const sections: { id: Section; label: string; icon: AssetIconName }[] = [
  { id: "chat", label: "Chat", icon: "message-square" },
  { id: "jira", label: "Jira", icon: "square-kanban" },
  { id: "sales", label: "Sales Info", icon: "briefcase-business" },
];

type SidebarProps = {
  open: boolean;
  onClose: () => void;
  onNewChat: () => void;
  newChatDisabled: boolean;
  section: Section;
  onSelectSection: (section: Section) => void;
  footer: ReactNode;
};

export default function Sidebar({ open, onClose, onNewChat, newChatDisabled, section, onSelectSection, footer }: SidebarProps) {
  return (
    <>
      {/* Mobile: the drawer overlays the chat, so give it a dismissible backdrop. */}
      <button
        type="button"
        aria-label="Close sidebar"
        tabIndex={-1}
        onClick={onClose}
        className={`sidebar-backdrop ${open ? "sidebar-backdrop--visible" : ""}`}
      />

      <aside
        id="app-sidebar"
        aria-label="Sidebar"
        inert={!open}
        className={`sidebar ${open ? "sidebar--open" : ""}`}
      >
        {/* Fixed inner width keeps labels from reflowing while the drawer animates. */}
        <div className="sidebar__inner">
          <div className="sidebar__header">
            <span className="sidebar__logo" aria-hidden="true">D</span>
            <span className="sidebar__brand">DCT</span>
            <button
              type="button"
              onClick={onClose}
              aria-label="Collapse sidebar"
              aria-controls="app-sidebar"
              aria-expanded={open}
              title="Collapse sidebar"
              className="sidebar__toggle"
            >
              <AssetIcon name="panel-left-close" size={18} />
            </button>
          </div>

          <div className="sidebar__cta">
            <button type="button" onClick={onNewChat} disabled={newChatDisabled} className="new-chat-button">
              <AssetIcon name="plus" size={17} />
              New Chat
            </button>
          </div>

          <nav aria-labelledby="features-label" className="sidebar__nav">
            <p id="features-label" className="sidebar__label">Features</p>
            <ul className="sidebar__list">
              {sections.map(({ id, label, icon }) => {
                const active = id === section;
                return (
                  <li key={id}>
                    <button
                      type="button"
                      onClick={() => onSelectSection(id)}
                      aria-current={active ? "page" : undefined}
                      className={`nav-item ${active ? "nav-item--active" : ""}`}
                    >
                      <AssetIcon name={icon} size={18} />
                      {label}
                    </button>
                  </li>
                );
              })}
            </ul>
          </nav>

          <div className="sidebar__footer">{footer}</div>
        </div>
      </aside>
    </>
  );
}
