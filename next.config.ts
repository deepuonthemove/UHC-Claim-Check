import type { NextConfig } from 'next';

const nextConfig: NextConfig = {
  poweredByHeader: false,
  serverExternalPackages: ['playwright-core', '@sparticuz/chromium'],
  outputFileTracingIncludes: {
    '/api/process': ['node_modules/@sparticuz/chromium/bin/**/*'],
  },
};

export default nextConfig;
