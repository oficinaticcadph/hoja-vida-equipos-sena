import type { Metadata } from "next";
import "./styles.css";

export const metadata: Metadata = {
  title: "Hoja de Vida Equipos SENA",
  description: "Generación de hojas de vida de equipos a partir del inventario SENA."
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="es">
      <body>{children}</body>
    </html>
  );
}
