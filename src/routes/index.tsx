import { createFileRoute } from "@tanstack/react-router";
import { FolioApp } from "@/components/folio-app";

export const Route = createFileRoute("/")({ component: Home });

function Home() {
  return <FolioApp />;
}
