import { Analytics } from "@vercel/analytics/next";

export const metadata = {
  title: "애슐리 인건비 모니터",
  description: "주간 인건비 모니터링 대시보드",
};

export default function RootLayout({ children }) {
  return (
    <html lang="ko">
      <body style={{ margin: 0 }}>
        {children}
        <Analytics />
      </body>
    </html>
  );
}
