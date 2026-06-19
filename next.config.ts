import type { NextConfig } from 'next';

const nextConfig: NextConfig = {
  poweredByHeader: false,
  serverExternalPackages: ['playwright-core', '@sparticuz/chromium'],
  outputFileTracingIncludes: {
    '/api/**/*': ['node_modules/@sparticuz/chromium/bin/**/*'],
  },
};

export default nextConfig;
