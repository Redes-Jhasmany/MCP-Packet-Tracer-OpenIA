# Codex MCP Packet Tracer

Repositorio de respaldo para la configuracion de Codex con Cisco Packet Tracer usando el servidor MCP completo de Packet Tracer.

## Contenido

- `docs/codex-mcp-packet-tracer-pasos.docx`: documento Word con los pasos completos para conectar Codex con Packet Tracer MCP.
- `scripts/create_codex_packet_tracer_doc.py`: script que genera el documento DOCX desde cero.
- `LICENSE`: licencia del repositorio.

## Resumen del flujo

El objetivo es que Codex pueda controlar Packet Tracer mediante herramientas MCP `pt_*`, por ejemplo:

- `pt_bridge_status`
- `pt_list_devices`
- `pt_full_build`
- `pt_query_topology`

El flujo final es:

```text
Usuario -> Codex -> Packet Tracer MCP -> MCP BUILDER -> Cisco Packet Tracer
```

## Instalacion usada

Servidor MCP completo:

```bash
git clone https://github.com/Mats2208/MCP-Packet-Tracer.git /home/jhasmany/Proyectos/CCNA/MCP-Packet-Tracer
cd /home/jhasmany/Proyectos/CCNA/MCP-Packet-Tracer
python3 -m venv .venv
.venv/bin/python -m pip install --upgrade pip
.venv/bin/pip install -e .
```

Registro en Codex:

```bash
codex mcp remove packet-tracer
codex mcp add packet-tracer -- /home/jhasmany/Proyectos/CCNA/MCP-Packet-Tracer/.venv/bin/python -m packet_tracer_mcp --stdio
codex mcp get packet-tracer
```

Extension de Packet Tracer:

```bash
mkdir -p /home/jhasmany/Proyectos/CCNA/MCP-Packet-Tracer/releases
curl -L --fail -o /home/jhasmany/Proyectos/CCNA/MCP-Packet-Tracer/releases/V5.2.pts https://github.com/Mats2208/MCP-Packet-Tracer/releases/download/v0.9.0/V5.2.pts
```

Luego, dentro de Packet Tracer:

1. Abrir `Extensions > Scripting > Configure PT Script Modules`.
2. Agregar `V5.2.pts`.
3. Seleccionar `MCP BUILDER` o `MCP Control Center`.
4. Pulsar `Start`.
5. Abrir `Extensions > MCP BUILDER` y dejar la ventana abierta.

## Verificacion

En una tarea nueva de Codex, ejecutar:

```text
pt_bridge_status
```

Resultado esperado:

```text
CONNECTED por HTTP (ventana MCP Control Center abierta) -- http://127.0.0.1:54321
```

Prueba real ejecutada:

```text
pt_list_devices
pt_full_build routers=1 pcs_per_lan=2 switches_per_router=1 dhcp=true routing=none router_model=2911 switch_model=2960-24TT template=single_lan deploy=true
pt_query_topology
```

Resultado comprobado:

- Topologia creada: `R1`, `SW1`, `PC1`, `PC2`.
- Dispositivos verificados: `4/4`.
- Enlaces verificados: `3/3`.
- Bridge conectado por HTTP en `127.0.0.1:54321`.

## Regenerar el documento

El script usa `python-docx`. Si estas trabajando desde Codex Desktop, puedes usar el Python del runtime de documentos. Ejemplo:

```bash
/home/jhasmany/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3 scripts/create_codex_packet_tracer_doc.py
```

El archivo se genera en:

```text
docs/codex-mcp-packet-tracer-pasos.docx
```

## Nota

Si Codex solo muestra herramientas `packet_tracer_*` y no las herramientas `pt_*`, la sesion actual cargo el MCP antiguo. Abre una tarea nueva de Codex o reinicia la sesion para que tome la configuracion actualizada.
