import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'UHC Claim Status Automation',
  description: 'Automate UHC Provider Portal claim status checks using Playwright and TOTP 2FA.',
};

export default function RootLayout({
  children,
}: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="en">
      {/* suppressHydrationWarning prevents false errors from browser extensions
          (password managers, etc.) that inject attributes into <body> before React loads */}
      <body suppressHydrationWarning>{children}</body>
    </html>
  );
}
