This is a [Next.js](https://nextjs.org) project bootstrapped with [`create-next-app`](https://nextjs.org/docs/app/api-reference/cli/create-next-app).

## Getting Started

First, run the development server:

```bash
npm run dev
# or
yarn dev
# or
pnpm dev
# or
bun dev
```

Open [http://localhost:3000](http://localhost:3000) with your browser to see the result.

You can start editing the page by modifying `app/page.tsx`. The page auto-updates as you edit the file.

This project uses [`next/font`](https://nextjs.org/docs/app/building-your-application/optimizing/fonts) to automatically optimize and load [Geist](https://vercel.com/font), a new font family for Vercel.

## Learn More

To learn more about Next.js, take a look at the following resources:


You can check out [the Next.js GitHub repository](https://github.com/vercel/next.js) - your feedback and contributions are welcome!

## Deploy on Vercel

The easiest way to deploy your Next.js app is to use the [Vercel Platform](https://vercel.com/new?utm_medium=default-template&filter=next.js&utm_source=create-next-app&utm_campaign=create-next-app-readme) from the creators of Next.js.
Check out the [Next.js deployment documentation](https://nextjs.org/docs/app/building-your-application/deploying) for more details.

## Deployment Notes

This app needs a separate Python backend for Excel parsing and HWP generation.

Set these environment variables in Vercel or your hosting platform:

- `BACKEND_API_URL`: public URL of the FastAPI backend, for example `https://your-backend.example.com`
- `FRONTEND_ORIGIN`: optional exact frontend origin for CORS, for example `https://your-app.vercel.app`
- `FRONTEND_ORIGIN_REGEX`: optional regex for allowed frontend origins, defaults to localhost and Vercel preview domains

The backend must run on a Windows machine if you use the current HWP automation flow, because it depends on `pywin32` and Hancom HWP COM automation.

## Easier Startup

On Windows, you can start the backend and public tunnel together by double-clicking `start-autoreport.bat` from the project root.

That launcher opens two windows:

- FastAPI on `http://127.0.0.1:8000`
- localtunnel on `https://autoreport-backend.loca.lt`

Keep both windows open while people are using the site.
Check out our [Next.js deployment documentation](https://nextjs.org/docs/app/building-your-application/deploying) for more details.
