# Hoja de Vida Equipos SENA

Aplicación web para generar hojas de vida de equipos usando el archivo de inventario y la plantilla oficial GTI-F-132.

## Qué hace

- Lee las pestañas **Funcionarios**, **FORMACION** y **ELECTRONICOS** del inventario.
- Busca por **Placa SENA** y permite filtrar por hoja de origen.
- Completa componente, marca, modelo, serial, placa, disco duro, procesador, RAM, cuentadante, sede, hostname y dominio.
- Permite corregir los datos antes de exportar.
- Descarga `HOJA_DE_VIDA_[PLACA].xlsx` editable.

## Ejecutar en el computador

1. Instale Node.js 18 o superior.
2. Abra una terminal en esta carpeta y ejecute:

   ```bash
   npm install
   npm run dev
   ```

3. Abra `http://localhost:3000` en el navegador.
4. Cargue primero `INVENTARIO SENA 2026.xlsx` y después `GTI-F-132 HOJA DE VIDA.xlsx`.

## Publicar en Vercel

1. Suba esta carpeta a un repositorio de GitHub.
2. En [Vercel](https://vercel.com), seleccione **Add New → Project** y elija el repositorio.
3. Vercel detectará Next.js. Presione **Deploy** sin cambiar la configuración.

Los archivos se leen en el navegador. No se almacenan de forma permanente en la aplicación.

## Mapeo configurado

La plantilla GTI-F-132 usa las siguientes celdas: `A23` componente, `B23` marca, `D23` modelo, `E23` serial, `G23` placa, `C17` cuentadante, `C19` sede, `J12` hostname, `J13` dominio, `E35` disco duro, `E36` procesador y `E37` memoria RAM.
