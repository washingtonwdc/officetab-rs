# API pública

Esta API é Windows-only. As estruturas persistidas derivam `Serialize` e `Deserialize`; use os métodos de importação para validar e migrar dados externos.

## Configuração

`persistence::OfficeTabConfig` contém as opções gerais, o tema, atalhos, favoritos, documentos recentes e grupos de abas. `OfficeTabConfig::default()` cria os padrões atuais.

```rust
use officetab_rs::persistence::{
    OfficeTabConfig, export_config_to_file, export_favorites_to_file,
    import_config_from_file, import_favorites_from_file,
};

let mut config = OfficeTabConfig::default();
config.add_favorite("Relatório", r"C:\Docs\Relatorio.xlsx");
config.add_recent_document("Relatório", r"C:\Docs\Relatorio.xlsx");
config.save_tab_group("Fechamento", vec![], Some("Excel"));

export_config_to_file(&config, std::path::Path::new("config.json"))?;
export_favorites_to_file(&config, std::path::Path::new("favorites.json"))?;

let mut restored = import_config_from_file(std::path::Path::new("config.json"))?;
import_favorites_from_file(&mut restored, std::path::Path::new("favorites.json"))?;
```

`load_effective_config` e `save_effective_config` escolhem o JSON portátil ao lado do módulo se ele existir; caso contrário, usam HKCU. `load_config` e `save_config` acessam o Registro diretamente. `migrate_config` valida JSON e migra esquemas anteriores.

## Temas

`theme::Theme::presets()` lista os presets disponíveis; `find_preset(name)` procura um tema pelo nome e `for_app(OfficeApp)` retorna a paleta recomendada por produto. `Theme::to_colorref` converte RGB para a ordem de bytes usada pelo GDI.

Para ativar as paletas por aplicativo, defina `config.general.adaptive_theme_by_app = true` e salve a configuração. O add-in aplica as cores base, selecionada e hover por custom draw do controle de abas.

## Add-in e janelas

`implementation::OfficeTab::new()` cria o orquestrador. `initialize()` inicia a detecção de janelas; `shutdown()` remove hook, timer, subclasses e controles. As operações Win32 de janela devem ser chamadas na thread proprietária do Office.

`tab_manager::obter_tipo_app` converte classes de janela conhecidas em `TipoAppOffice`. `listar_janelas_processo` enumera apenas janelas visíveis do processo atual. `OFFICETAB_TIMER_ID` identifica o timer de cada janela monitorada.

## Persistência de licença

`license` é experimental e não participa da ativação do add-in. A chave de validação embutida no código não é privada e o hash auxiliar não é assinatura. Não use essa API para licenças comerciais ou para proteger dados.

## Resultados de erro

APIs Win32 retornam `windows::core::Result`. Configuração inválida e JSON incompatível retornam erro antes de serem aplicados. Chamadas COM exportadas convertem falhas internas em HRESULT para preservar o ABI do Office.
