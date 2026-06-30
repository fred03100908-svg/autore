import type { Metadata } from "next";
import { Bebas_Neue } from "next/font/google";
import "./globals.css";

const displayFont = Bebas_Neue({
  subsets: ["latin"],
  weight: "400",
  variable: "--font-display",
});

export const metadata: Metadata = {
  title: "AutoReport",
  description: "Excel과 HWP 템플릿으로 보고서를 자동 생성하는 웹서비스",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="ko">
      <body className={displayFont.variable}>{children}</body>
    </html>
  );
}