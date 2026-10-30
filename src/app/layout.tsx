import type { Metadata } from 'next';

export const metadata: Metadata = {
  title: 'PPT Creation Photo Upload - Multi Zone Platform',
  description: 'North Zone and East Zone site photo packs in one platform.',
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
