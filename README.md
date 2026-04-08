# Penta Export Automation

Automatización robusta con **Playwright + Node.js** para extraer datos de:

- Importadores en Argentina
- Exportadores de Argentina

Genera salida consolidada en:

- `data/importadores_argentina.xlsx`
- `data/exportadores_argentina.xlsx`
- `data/importadores_argentina.json`
- `data/exportadores_argentina.json`

## Requisitos

- Node.js 18+ (recomendado)

## Configuración

1. Instalar dependencias:

```bash
npm install
```

2. Crear `.env` desde el ejemplo:

```bash
cp .env.example .env
```

3. Completar variables en `.env`:

```env
PENTA_USER=TU_USUARIO
PENTA_PASS=TU_PASSWORD
PENTA_BASE_URL=https://app.penta-transaction.com
```

## Ejecución

```bash
node scripts/penta-export.js
```

También disponible:

```bash
npm run export:penta
```
