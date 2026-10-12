# Guia Completo dos Instaladores — Office Tab (Rust) 📑🚀

Este guia detalha todos os instaladores e pacotes disponibilizados para o **OfficeTab-rs**, permitindo escolher a opção mais conveniente para seu perfil (usuário final, administrador ou desenvolvedor).

---

## 📂 Onde estão localizados os instaladores?

Os instaladores foram compilados e organizados nas seguintes pastas para acesso imediato:

1. **No Repositório Principal (`C:\Users\washi\officetab-rs\dist\`):**
   - `OfficeTab-rs-Setup-v0.1.0.exe` — Instalador Gráfico Oficial Windows (Inno Setup)
   - `OfficeTab-rs-v0.1.0-Portable.zip` — Pacote Portátil Completo (ZIP)
   - `original\SetupOfficeTabPro.exe` — Instalador comercial original analisado (ExtendOffice v19)

2. **Na Pasta de Entrega do Workspace (`default-workspace\installers\`):**
   - `OfficeTab-rs-Setup-v0.1.0.exe` — Instalador Executável
   - `OfficeTab-rs-v0.1.0-Portable.zip` — Pacote ZIP Portátil
   - `SetupOfficeTabPro.exe` — Instalador original de referência
   - `Instalar.bat` / `Desinstalar.bat` / `Testar.bat` / `Configurar.bat` — Scripts de 1-clique
   - `officetab_rs.dll` / `officetab-config.exe` — Binários compilados diretos

---

## 🎯 Escolha o Seu Método de Instalação

### 1. 🌟 Instalador Gráfico Windows (`OfficeTab-rs-Setup-v0.1.0.exe`) — Recomendado
- **Tamanho:** ~2.85 MB (compactação ultra LZMA2).
- **Interface:** Assistente visual amigável nos idiomas Português (Brasil) e Inglês.
- **O que ele faz:**
  - Copia a DLL COM de alta performance (`officetab_rs.dll`) e o painel de configurações (`officetab-config.exe`).
  - Registra automaticamente o ProgID `OfficeTabRs.Connect` e CLSID `{9a1929ef-1495-413e-993c-326dd15e35df}` no Registro do Windows (`HKCU`).
  - Habilita o suplemento com `LoadBehavior = 3` nos **7 aplicativos Office** (Word, Excel, PowerPoint, Access, Project, Visio e Publisher).
  - Limpa entradas na lista de resiliência (*DisabledItems*) do Office para evitar que seja bloqueado.
  - Cria atalhos no Menu Iniciar para "Painel de Configurações", "Diagnóstico" e "Desinstalar".
  - Registra o desinstalador oficial no **Painel de Controle / Configurações > Aplicativos Instalados do Windows**.

### 2. ⚡ Instalador de 1-Clique (`Instalar.bat`)
- **Como usar:** Basta dar dois cliques em `Instalar.bat`.
- **Vantagem:** Não requer abrir janelas de terminal ou configurar permissões de política de execução do PowerShell (`ExecutionPolicy Bypass` embutido).
- **Desinstalação:** Dê dois cliques em `Desinstalar.bat` para remover todos os registros limpos.

### 3. 🧰 Pacote Portátil (`OfficeTab-rs-v0.1.0-Portable.zip`)
- **Tamanho:** ~1.66 MB.
- **Ideal para:** Pendrives, ambientes corporativos restritos ou testes isolados.
- **Como usar:**
  1. Extraia o arquivo ZIP em qualquer pasta desejada.
  2. Dê duplo clique em `Instalar.bat`.
  3. Dê duplo clique em `Configurar.bat` para ajustar cores e atalhos.

### 4. 💻 Scripts PowerShell para Automação / CI/CD
- `Instalar-OfficeTab.ps1`: Aceita parâmetro opcional `-DllPath <caminho>` para automatizar deploys.
- `Desinstalar-OfficeTab.ps1`: Limpeza total idempotente do registro HKCU.
- `Testar-OfficeTab.ps1`: Executa verificação completa de saúde do sistema (integridade da DLL, chaves dos 7 apps, processos ativos e logs).

### 5. 🔍 Instalador Original de Referência (`SetupOfficeTabPro.exe`)
- Localizado em `installers\SetupOfficeTabPro.exe` e `C:\Users\washi\Downloads\SetupOfficeTabPro.exe`.
- Fornecido para fins de comparação com a versão original comercial durante testes de compatibilidade.

---

## 🛠️ Verificação e Diagnóstico Pós-Instalação

Após a instalação, abra qualquer aplicativo do Microsoft Office (ex: Word ou Excel). A barra de abas moderna será exibida automaticamente logo abaixo da faixa de opções / barra de fórmulas.

Para testar ou alterar configurações:
- Execute `Configurar.bat` ou abra **"Painel de Configurações do Office Tab"** no Menu Iniciar.
- Para verificar a integridade da instalação a qualquer momento, execute `Testar.bat`.
