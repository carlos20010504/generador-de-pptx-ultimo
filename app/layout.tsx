import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Socya PPTX Generator",
  description: "Organiza archivos Excel y genera presentaciones PowerPoint listas para negocio.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="es" className="h-full antialiased">
      <body className="min-h-full flex flex-col">{children}</body>
    </html>
  );
}
