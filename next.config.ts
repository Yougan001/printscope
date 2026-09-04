import type { NextConfig } from 'next';

const nextConfig: NextConfig = {
  output: 'export',
  trailingSlash: true,
  basePath: process.env.GITHUB_PAGES ? '/printscope' : undefined,
};

export default nextConfig;
