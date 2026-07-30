import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "物件管理總表｜單機版",
  description: "房仲物件委託、封存、鑰匙與前台總表管理",
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="zh-Hant"><body>{children}</body></html>;
}
