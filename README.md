# SigaExcel

SigaExcel e uma versao atualizada do antigo executavel Rezende.exe usado para interpretar scripts texto e automatizar o Microsoft Excel.

O projeto nasceu de uma necessidade pratica: manter funcionando os relatorios gerados por scripts legados depois que atualizacoes recentes do Microsoft Office/Windows passaram a afetar a automacao Excel/OLE usada pelo executavel antigo. O problema mais visivel apareceu em comandos `Escreve` com parametros de cor de borda e cor de fundo, que antes eram aceitos pelo Excel e passaram a falhar em alguns ambientes.

A proposta do SigaExcel e preservar a compatibilidade com os scripts existentes, corrigir os pontos frageis da automacao COM/OLE e adicionar modos de execucao mais robustos, sem exigir mudanca imediata nos arquivos `.txt` ja utilizados.

## Agradecimentos
Enquanto eu fazia a engenharia reversa do rezende.exe, descobri que o script foi criado em 2005 e vem funcionando em diversas empresas por onde passei.

Um agradecimento ao Rafael Rezende, que criou e deixou disponível a versão anterior. Infelizmente essa versão foi afetada pela Microsoft e precisou ser ajustada.

Meu objetivo com esse update foi preservar tudo o que foi feito no script do Rezende, mas deixando aberto para futuras evoluções.

Por um momento pensei em deixar isso só comigo, mas enquanto trabalhava eu fiquei pensando em quantos relatórios em excel eu criei utilizando o rezende.exe...

Então resolvi criar esse git público para essa versão nova. Espero que ajude alguém.

Bom... Vamos lá!

## Identificacao

```text
Produto: SigaExcel
Versao: v1
Responsavel: Victor Cardoso
Contato: vhsbc92@gmail.com
Data do snapshot: 2026-06-18
```

## O que esta versao resolve

- Executa scripts legados que automatizam o Excel.
- Aceita delimitadores `;` e `|`.
- Mantem compatibilidade com caminhos absolutos antigos.
- Aceita numeros com zeros a esquerda, como `001` e `02`.
- Corrige a aplicacao de `corBorda` e `corFundo` usando conversao numerica explicita para `ColorIndex`.
- Exibe erros em janela (`MessageBox`), sem abrir uma janela de CMD no uso normal.
- Inclui modo rapido com Excel invisivel durante a geracao e overlay de progresso.
- Inclui modo moderno experimental para aplicar uma apresentacao visual mais atual ao relatorio.

## Requisitos

- Windows.
- Microsoft Excel instalado e registrado como `Excel.Application`.
- Permissao para abrir os modelos e gravar os arquivos de saida usados pelos scripts.
- Scripts compativeis com o formato legado do SigaExcel/rezende.

## Estrutura do pacote

```text
sigaexcel.exe        Executavel pronto para uso
src/Program.cs       Codigo-fonte correspondente
docs/sigaexcel.md    Documentacao tecnica detalhada
docs/logo.png        Imagem original do icone
assets/sigaexcel.ico Icone embutido no executavel
SHA256SUMS.txt       Hashes do snapshot
```

## Como usar

Execute o `sigaexcel.exe` informando o caminho do script:

```powershell
.\sigaexcel.exe C:\sigaexcel\scripts\20260617.txt
```

O script deve conter comandos como:

```text
AbreExcel;S;
AbreArquivo|C:\SigaExcel\Modelos\Export_pc.xlsx|
Escreve;006;001;VICTOR CARDOSO;1;Arial;08;S;N;N;10;S;10;20;
GravaArquivo|C:\SigaExcel\Planilhas\Relatorio.xlsx|
```

Durante a execucao, o SigaExcel interpreta os comandos em sequencia e aplica as alteracoes no Excel via automacao COM.

## Opcoes de linha de comando

### Ajuda

```powershell
.\sigaexcel.exe --help
```

Mostra a tela de uso do programa.

### Validar leitura do script sem abrir Excel

```powershell
.\sigaexcel.exe --dry-run C:\sigaexcel\scripts\20260617.txt
```

Usa o parser do SigaExcel e lista os comandos encontrados, sem executar automacao no Excel. E util para conferir se o arquivo esta sendo interpretado corretamente.

### Modo rapido

```powershell
.\sigaexcel.exe --fast C:\sigaexcel\scripts\007124.txt
```

Executa o relatorio com ajustes de performance:

- Excel invisivel durante a geracao.
- Overlay de progresso sempre visivel.
- `ScreenUpdating` desativado.
- `EnableEvents` desativado.
- `DisplayAlerts` desativado.
- Calculo em modo manual durante a execucao.
- Restauracao das configuracoes ao final.
- Excel visivel novamente no final.

### Modo moderno experimental

```powershell
.\sigaexcel.exe --modern C:\sigaexcel\scripts\20260617.txt
```

Executa o script legado normalmente e, ao final, aplica uma formatacao visual moderna no workbook.

Tambem e possivel informar um arquivo de saida:

```powershell
.\sigaexcel.exe --modern --output C:\sigaexcel\planilhas\SigaExcelModern.xlsx C:\sigaexcel\scripts\20260617.txt
```

### Modo rapido + moderno

```powershell
.\sigaexcel.exe --fast --modern C:\sigaexcel\scripts\20260617.txt
```

Combina a execucao acelerada com a formatacao moderna ao final.

### Testar tela de erro

```powershell
.\sigaexcel.exe --test-error
```

Exibe uma janela de erro de teste. Serve para validar que as mensagens aparecem corretamente no ambiente do usuario.

## Comandos de script suportados

Esta versao implementa os principais comandos usados nos scripts analisados:

```text
AbreExcel
AbreArquivo
AdicionaSheet
Escreve
EscreveFormula
MergeCelulas
FormataNumeroCelula
AutoFormata
ConfiguraPagina
RenomeiaSheet
GravaArquivo
FechaExcel
Imprime
VisualizaImpressao
```

Os comandos sao tratados sem diferenciar maiusculas e minusculas.

## Observacao sobre cores

Nos scripts legados, os parametros de cor normalmente aparecem como indices da paleta do Excel:

```text
10;02
10;15
```

O SigaExcel converte esses valores explicitamente para inteiro antes de chamar o Excel:

```text
10 -> ColorIndex 10
02 -> ColorIndex 2
15 -> ColorIndex 15
```

Essa regra evita depender da conversao implicita do COM/OLE, que foi um dos pontos suspeitos na falha observada no executavel legado.

## Verificacao do executavel

Hash SHA-256 do `sigaexcel.exe` deste snapshot:

```text
88CE7000560258CC1CB5A84DBEEDC3B13C2EEEB70CFA61F12AA75C6253DAE0F9
```

Para conferir localmente:

```powershell
Get-FileHash .\sigaexcel.exe -Algorithm SHA256
```

## Documentacao tecnica

A documentacao completa da analise, comportamento legado, comandos conhecidos, hipoteses tecnicas e validacoes fica em:

```text
docs/sigaexcel.md
```
