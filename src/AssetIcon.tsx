import type { CSSProperties } from "react";

export type AssetIconName =
  | "arrow-up"
  | "briefcase-business"
  | "building-2"
  | "circle-check"
  | "clock"
  | "contact-round"
  | "file-text"
  | "list-todo"
  | "mail"
  | "message-square"
  | "panel-left-close"
  | "panel-left-open"
  | "plus"
  | "square-kanban"
  | "square-pen"
  | "sun"
  | "moon"
  | "user-round"
  | "users"
  | "zap";

type AssetIconProps = {
  name: AssetIconName;
  size?: number;
  className?: string;
};

export default function AssetIcon({ name, size, className = "" }: AssetIconProps) {
  const style = size ? { "--asset-icon-size": `${size}px` } as CSSProperties : undefined;
  const common = {
    className: `asset-icon ${className}`.trim(),
    width: size,
    height: size,
    alt: "",
    "aria-hidden": true,
  } as const;

  return (
    <>
      <img {...common} style={style} src={`/assets/${name}.light.svg`} data-theme-icon="light" />
      <img {...common} style={style} src={`/assets/${name}.dark.svg`} data-theme-icon="dark" />
    </>
  );
}
