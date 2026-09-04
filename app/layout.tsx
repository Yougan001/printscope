import type { Metadata } from 'next';
import './globals.css';
export const metadata: Metadata = {
  metadataBase: new URL('https://yougan001.github.io/printscope/'),
  title: 'Printscope — Excel blank-page diagnostics',
  description:
    'Inspect workbook print areas and distant cells to find likely causes of unexpected printed pages.',
  alternates: { canonical: 'https://yougan001.github.io/printscope/' },
  icons: {
    icon: process.env.GITHUB_PAGES ? '/printscope/favicon.svg' : '/favicon.svg',
  },
  openGraph: {
    title: 'Printscope — Excel blank-page diagnostics',
    description:
      'Read-only XLSX print-area and hidden-cell diagnostics, entirely in your browser.',
    url: 'https://yougan001.github.io/printscope/',
    type: 'website',
  },
};
export default function RootLayout({
  children,
}: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
