import { createFileRoute, useNavigate } from "@tanstack/react-router";
import { useAppState } from "@/lib/app-context";
import { AppShell } from "@/components/xls-vcard/AppShell";
import { DropStep } from "@/components/xls-vcard/XlsToVcardApp";

export const Route = createFileRoute("/")({
  head: () => ({
    meta: [
      { title: "XLS → vCard — convert spreadsheets to contact files" },
      {
        name: "description",
        content:
          "Drop an Excel or CSV file, map columns to vCard fields with full label control and filters, download a .vcf. Runs entirely in your browser.",
      },
      { property: "og:title", content: "XLS → vCard" },
      {
        property: "og:description",
        content: "Convert spreadsheets to vCard contacts in your browser.",
      },
      { property: "og:type", content: "website" },
      { name: "twitter:card", content: "summary" },
    ],
  }),
  component: Index,
});

function Index() {
  const { handleFile } = useAppState();
  const navigate = useNavigate();

  return (
    <AppShell>
      <DropStep
        onFile={(f) => {
          handleFile(f).then(() => navigate({ to: "/preview" }));
        }}
      />
    </AppShell>
  );
}
