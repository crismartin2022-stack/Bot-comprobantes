# 🤖 Bot Comprobantes Agilpagos

Bot de Telegram que analiza comprobantes bancarios con IA y exporta a Excel semanal.

---

## ⚙️ Configuración paso a paso

### 1. Crear el bot en Telegram
1. Abrí Telegram y buscá **@BotFather**
2. Enviá `/newbot`
3. Elegí un nombre (ej: `Agilpagos Comprobantes`)
4. Elegí un username (ej: `agilpagos_bot`)
5. Guardá el **TOKEN** que te da BotFather

### 2. Agregar el bot al grupo
1. Abrí el grupo de Telegram
2. Tocá el nombre del grupo → Agregar miembros
3. Buscá tu bot por username y agregalo
4. (Opcional) Dale permisos de administrador para que pueda leer mensajes

### 3. Obtener el ID del grupo (opcional pero recomendado)
1. Agregá @userinfobot al grupo
2. Escribí cualquier mensaje
3. Te dará el `Chat ID` del grupo (número negativo, ej: `-1001234567890`)
4. Guardalo para el paso siguiente

### 4. Obtener API Key de Anthropic
1. Entrá a https://console.anthropic.com
2. Creá una cuenta o iniciá sesión
3. Andá a **API Keys** → **Create Key**
4. Guardá la key

### 5. Deploy en Railway
1. Creá cuenta en https://railway.app (gratis, sin tarjeta)
2. Creá un **New Project** → **Deploy from GitHub repo**
   - O usá **Deploy from local** subiendo esta carpeta
3. En el proyecto, andá a **Variables** y agregá:

   | Variable | Valor |
   |---|---|
   | `TELEGRAM_TOKEN` | El token de BotFather |
   | `ANTHROPIC_API_KEY` | Tu API key de Anthropic |
   | `ALLOWED_GROUP_ID` | El ID del grupo (ej: `-1001234567890`) — opcional |

4. Railway detecta el `Procfile` y arranca el bot automáticamente

---

## 📱 Comandos del bot

| Comando | Función |
|---|---|
| `/start` | Bienvenida e instrucciones |
| `/resumen` | Ver cuántos comprobantes hay cargados |
| `/excel` | Generar y descargar el Excel semanal |
| `/nueva_semana` | Iniciar semana nueva (borra registros actuales) |
| `/borrar` | Borrar todos los registros |
| `/ayuda` | Ayuda |

---

## 🔄 Flujo de uso

1. Alguien sube una foto de comprobante al grupo
2. El bot responde automáticamente con los datos extraídos
3. Si falta el CVU, avisa con 🔴 y muestra botón para corregir
4. Al final de la semana: `/excel` → descargás el archivo
5. `/nueva_semana` para empezar la siguiente semana

---

## 📊 Formato del Excel

El Excel exportado tiene el mismo formato que Agilpagos:
- `FECHA DE ENVIO`
- `TRF O DEPOSITO`
- `TITULAR DE LA CTA`
- `FECHA TICKET`
- `HORA TICKET`
- `CUENTA` (últimos 4 dígitos del CVU — en rojo si falta)
- `MONTO`
- Y columnas adicionales de contexto

---

## ⚠️ Notas importantes

- Los datos se guardan **en memoria** mientras el bot corre. Si Railway reinicia el servicio se pierden.
- Exportá el Excel antes de usar `/nueva_semana`
- El plan gratuito de Railway tiene 500 horas/mes — suficiente para uso normal
