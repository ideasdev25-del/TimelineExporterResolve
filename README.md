# DaVinci Resolve Timeline Exporter

Ferramenta nativa para DaVinci Resolve que gera relatórios detalhados da Timeline em formato HTML, similar à versão para Adobe Premiere Pro.

## Como Instalar

1. Abra o DaVinci Resolve.
2. No menu superior, vá em **Help > Documentation > Developer**.
3. Isso abrirá a pasta de scripts. Navegue até:
   - `Support/Fusion/Scripts/Utility`
4. Copie o arquivo `TimelineExporter.py` para esta pasta.

## Como Usar

1. No DaVinci Resolve, abra a Timeline que deseja exportar.
2. Vá no menu **Workspace > Scripts > TimelineExporter**.
3. A interface (Fusion UI) irá abrir:
   - Cole os nomes dos arquivos que deseja filtrar (um por linha).
   - Escolha se quer analisar Vídeo, Áudio ou Ambos.
   - Filtre por cor de clipe se necessário.
   - Selecione as colunas que deseja no relatório.
   - Selecione a pasta de saída.
4. Clique em **GENERATE REPORT**.

## Requisitos

- DaVinci Resolve 17, 18 ou 19 (Funciona na versão Free).
- Python 3 instalado no sistema e configurado nas preferências do Resolve (**Preferences > System > General > External scripting using Python**).
