import { QueryClient, QueryClientProvider, useQueryClient } from "@tanstack/react-query";
import {
  Outlet,
  Link,
  createRootRouteWithContext,
  useRouter,
  HeadContent,
  Scripts,
  useNavigate,
} from "@tanstack/react-router";
import { useEffect, type ReactNode } from "react";

import appCss from "../styles.css?url";
import { reportAppError } from "../lib/app-error-reporting";
import { AppProvider } from "../lib/app-context";

function NotFoundComponent() {
  return (
    <div className="flex min-h-screen items-center justify-center bg-background px-4">
      <div className="max-w-md text-center">
        <h1 className="text-7xl font-bold text-foreground">404</h1>
        <h2 className="mt-4 text-xl font-semibold text-foreground">Page not found</h2>
        <p className="mt-2 text-sm text-muted-foreground">
          The page you're looking for doesn't exist or has been moved.
        </p>
        <div className="mt-6">
          <Link
            to="/"
            className="inline-flex items-center justify-center rounded-md bg-primary px-4 py-2 text-sm font-medium text-primary-foreground transition-colors hover:bg-primary/90"
          >
            Go home
          </Link>
        </div>
      </div>
    </div>
  );
}

function ErrorComponent({ error, reset }: { error: Error; reset: () => void }) {
  console.error(error);
  const router = useRouter();
  const queryClient = useQueryClient();
  const navigate = useNavigate();
  useEffect(() => {
    reportAppError(error, { boundary: "tanstack_root_error_component" });
  }, [error]);

  return (
    <div className="flex min-h-screen items-center justify-center bg-background px-4">
      <div className="max-w-md text-center">
        <h1 className="text-xl font-semibold tracking-tight text-foreground">
          This page didn't load
        </h1>
        <p className="mt-2 text-sm text-muted-foreground">
          Something went wrong on our end. You can try refreshing or head back home.
        </p>
        <div className="mt-6 flex flex-wrap justify-center gap-2">
          <button
            onClick={() => {
              router.invalidate();
              reset();
            }}
            className="inline-flex items-center justify-center rounded-md bg-primary px-4 py-2 text-sm font-medium text-primary-foreground transition-colors hover:bg-primary/90"
          >
            Try again
          </button>
          <button
            onClick={() => {
              if ("caches" in window) {
                caches.keys().then((names) => names.forEach((n) => caches.delete(n)));
              }
              window.location.reload();
            }}
            className="inline-flex items-center justify-center rounded-md bg-secondary px-4 py-2 text-sm font-medium text-secondary-foreground transition-colors hover:bg-secondary/80"
          >
            Clear cache only
          </button>
          <button
            onClick={() => {
              localStorage.clear();
              queryClient.clear();
              if ("caches" in window) {
                caches.keys().then((names) => names.forEach((n) => caches.delete(n)));
              }
              navigate({ to: "/" });
            }}
            className="inline-flex items-center justify-center rounded-md bg-destructive px-4 py-2 text-sm font-medium text-destructive-foreground transition-colors hover:bg-destructive/90"
          >
            Clear all data & restart
          </button>
          <a
            href="/"
            className="inline-flex items-center justify-center rounded-md border border-input bg-background px-4 py-2 text-sm font-medium text-foreground transition-colors hover:bg-accent"
          >
            Go home
          </a>
        </div>
      </div>
    </div>
  );
}

export const Route = createRootRouteWithContext<{ queryClient: QueryClient }>()({
  head: () => ({
    meta: [
      { charSet: "utf-8" },
      { name: "viewport", content: "width=device-width, initial-scale=1" },
      { title: "XLS → vCard — Convert spreadsheets to contact files" },
      { name: "description", content: "Free online tool to convert XLS, XLSX, and CSV files to vCard (.vcf) format. Map columns, filter rows, and download contacts. No uploads — everything stays in your browser." },
      { name: "author", content: "Mohamed Faaris" },
      { property: "og:title", content: "XLS → vCard — Convert spreadsheets to contact files" },
      { property: "og:description", content: "Free online tool to convert XLS, XLSX, and CSV files to vCard (.vcf) format. Map columns, filter rows, and download contacts. No uploads — everything stays in your browser." },
      { property: "og:url", content: "https://xls-to-vcard-magic.vercel.app" },
      { property: "og:type", content: "website" },
      { property: "og:image", content: "https://xls-to-vcard-magic.vercel.app/og-image.png" },
      { name: "twitter:card", content: "summary_large_image" },
      { name: "twitter:title", content: "XLS → vCard — Convert spreadsheets to contact files" },
      { name: "twitter:description", content: "Free online tool to convert XLS, XLSX, and CSV files to vCard (.vcf) format. Map columns, filter rows, and download contacts. No uploads — everything stays in your browser." },
      { name: "twitter:image", content: "https://xls-to-vcard-magic.vercel.app/og-image.png" },
      { name: "canonical", content: "https://xls-to-vcard-magic.vercel.app" },
    ],
    links: [
      {
        rel: "stylesheet",
        href: appCss,
      },
      { rel: "icon", href: "/favicon.ico", type: "image/x-icon" },
      { rel: "icon", href: "/favicon-32x32.png", sizes: "32x32", type: "image/png" },
      { rel: "icon", href: "/favicon-16x16.png", sizes: "16x16", type: "image/png" },
      { rel: "apple-touch-icon", href: "/apple-touch-icon.png" },
      { rel: "manifest", href: "/manifest.json" },
      { rel: "canonical", href: "https://xls-to-vcard-magic.vercel.app" },
    ],
  }),
  shellComponent: RootShell,
  component: RootComponent,
  notFoundComponent: NotFoundComponent,
  errorComponent: ErrorComponent,
});

function RootShell({ children }: { children: ReactNode }) {
  return (
    <html lang="en">
      <head>
        <HeadContent />
        <meta name="theme-color" content="#0a0a0a" />
      </head>
      <body>
        {children}
        <Scripts />
        <script dangerouslySetInnerHTML={{ __html: `'serviceWorker'in navigator&&navigator.serviceWorker.register('/sw.js')` }} />
      </body>
    </html>
  );
}

function RootComponent() {
  const { queryClient } = Route.useRouteContext();

  return (
    <QueryClientProvider client={queryClient}>
      <AppProvider>
        {/* Required: nested routes render here. Removing <Outlet /> breaks all child routes. */}
        <Outlet />
      </AppProvider>
    </QueryClientProvider>
  );
}
