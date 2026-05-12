import type { Metadata, Viewport } from "next";
import { Jost, Nunito_Sans } from "next/font/google";
import "./globals.css";
import { I18nProvider } from "@/utils/i18n";

// Socya: Futura para piezas gráficas → Jost en web (sustituto geométrico).
const jost = Jost({
  subsets: ["latin"],
  weight: ["400", "500", "600", "700", "800"],
  variable: "--font-jost",
  display: "swap",
});

// Socya: Calibri para documentos → Nunito Sans en web.
const nunitoSans = Nunito_Sans({
  subsets: ["latin"],
  weight: ["400", "500", "600", "700"],
  variable: "--font-nunito-sans",
  display: "swap",
});

export const metadata: Metadata = {
  title: "Socya PPTX Generator",
  description: "Organiza archivos Excel y genera presentaciones PowerPoint listas para negocio.",
};

export const viewport: Viewport = {
  width: "device-width",
  initialScale: 1,
  maximumScale: 5,
  themeColor: "#087062",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="es" className={`${jost.variable} ${nunitoSans.variable} antialiased`}>
      <body>
        <I18nProvider>{children}</I18nProvider>
      </body>
    </html>
  );
}
