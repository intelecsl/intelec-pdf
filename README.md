# Correo a PDF Intelec (v1.3)

Complemento de Outlook: botón **Guardar PDF** en la cinta al leer un correo → genera el PDF y lo descarga.

- Cabecera: asunto, De, Enviado, Para, CC y **lista de adjuntos** (nombre y tamaño).
- Cuerpo HTML con estilos e imágenes incrustadas (cid:) — texto seleccionable/buscable.
- Nombre: `AAAA-MM-DD_Remitente_Asunto.pdf` (sin tildes ni caracteres especiales; quita RE:/RV:/FW:).
- Pie con nombre y nº de página. Producer/Creator: "Intelec Correo a PDF v1.3".
- Descarga directa desde el panel; botón de respaldo «Abrir en ventana» (Dialog API) si un Outlook la bloquea.

## Publicación
Repo `intelecsl/intelec-pdf`, GitHub Pages desde `main` / raíz → `https://intelecsl.github.io/intelec-pdf/`.

## Instalación
Centro de administración de M365 → Aplicaciones integradas → Cargar aplicación personalizada → URL del manifest:
`https://intelecsl.github.io/intelec-pdf/manifest.xml`
(o en Outlook: Obtener complementos → Mis complementos → Agregar desde archivo).

## Regla de versiones
Con cada cambio subir versión en: badge HTML + `VERSION` en taskpane.js, `APP_VERSION` en pdfbuilder.js (PDF Producer),
`<Version>` de manifest.xml, `"version"` de manifest.json (y los `?v=` de caché).

GUID manifest: `ba51c3e1-1e94-4d40-a481-70e7420dae37` (no cambiar).

## Limitaciones
- Imágenes externas sin CORS pueden no aparecer (las incrustadas sí).
- Fuentes estándar PDF (Helvetica): emojis y alfabetos no latinos no se representan.
- `lib/jspdf.umd.min.js` (2.5.2) lleva un parche mínimo para fijar el Producer (`window.__PDF_PRODUCER__`).
