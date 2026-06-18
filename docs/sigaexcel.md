# Documentacao Tecnica do SigaExcel

## 1. Identificacao

```text
Produto: SigaExcel
Versao documentada: v1
Responsavel pela atualizacao: Victor Cardoso
Contato: vhsbc92@gmail.com
Data do snapshot: 2026-06-18
Executavel: sigaexcel.exe
Fonte principal: src/Program.cs
```

O SigaExcel e uma versao atualizada do executavel legado conhecido como `rezende.exe`, usado para interpretar scripts texto e automatizar o Microsoft Excel.

Esta documentacao descreve:

- o contexto que levou a nova implementacao;
- o formato dos scripts;
- os comandos suportados na v1;
- as opcoes de linha de comando;
- as decisoes de compatibilidade;
- os pontos tecnicos que ainda exigem validacao.

## 2. Contexto

O executavel legado automatizava o Excel por COM/OLE. O uso normal era executar um arquivo `.txt` contendo comandos sequenciais, por exemplo:

```powershell
C:\sigaexcel\rezende.exe C:\sigaexcel\scripts\20260617.txt
```

Apos atualizacoes do Microsoft Office/Windows liberadas em junho de 2026, alguns ambientes passaram a apresentar falha ao executar scripts que antes funcionavam. O caso mais evidente apareceu no comando `Escreve` quando os parametros extras de cor de borda e cor de fundo eram informados:

```text
Escreve;006;001;VICTOR CARDOSO;1;Arial;08;S;N;N;10;S;10;20;
```

Ao remover os dois ultimos parametros, o comando voltava a funcionar:

```text
Escreve;006;001;VICTOR CARDOSO;1;Arial;08;S;N;N;10;S;
```

A conclusao tecnica adotada para a v1 foi: preservar o formato legado, mas evitar depender de conversao implicita de tipos feita pelo COM/OLE. No SigaExcel, valores como `10`, `02` e `15` sao convertidos explicitamente para inteiros antes de serem aplicados como `ColorIndex` no Excel.

Referencias oficiais usadas durante a investigacao:

- Microsoft Support: June 2026 updates for Microsoft Office: https://support.microsoft.com/en-us/topic/june-2026-updates-for-microsoft-office-0a7978da-f46a-428c-9632-5b690a27e20a
- Microsoft Support: Description of the security update for Excel 2016: June 9, 2026 (KB5002877): https://support.microsoft.com/en-us/topic/description-of-the-security-update-for-excel-2016-june-9-2026-kb5002877-bca1344e-015d-46d9-a708-49cd74b75b74
- Microsoft Support: June 9, 2026 Windows cumulative update KB5094126: https://support.microsoft.com/en-us/topic/june-9-2026-kb5094126-os-builds-26200-8655-and-26100-8655-1a9bcba6-5f53-4075-8156-fe11ac631737
- Excel `Range.BorderAround`: https://learn.microsoft.com/en-us/office/vba/api/excel.range.borderaround
- Excel `Border.ColorIndex`: https://learn.microsoft.com/en-us/office/vba/api/excel.border.colorindex
- Excel `Borders.ColorIndex`: https://learn.microsoft.com/en-us/office/vba/api/excel.borders.colorindex
- Excel `Interior.ColorIndex`: https://learn.microsoft.com/en-us/office/vba/api/excel.interior.colorindex

## 3. Objetivo da v1

A v1 tem um objetivo conservador: substituir o executavel legado nos fluxos ja existentes sem exigir alteracao imediata nos scripts.

Escopo principal:

- interpretar scripts texto no formato historico;
- abrir o Excel por COM/OLE;
- executar comandos em sequencia;
- preservar delimitadores `;` e `|`;
- aceitar indices numericos com zeros a esquerda;
- corrigir a aplicacao de cor de borda e cor de fundo;
- exibir erros em janela para o usuario final;
- oferecer modo rapido para reduzir tempo de geracao;
- manter o Excel visivel ao final quando o modo rapido for usado.

Fora do escopo da v1:

- remover a dependencia do Excel instalado;
- reimplementar calculo de formulas fora do Excel;
- garantir compatibilidade total com todos os comandos existentes no binario legado;
- substituir integralmente macros, impressoes e particularidades de ambiente sem validacao caso a caso.

## 4. Requisitos de ambiente

```text
Sistema operacional: Windows
Aplicativo necessario: Microsoft Excel
Registro COM necessario: Excel.Application
Interface: execucao por linha de comando ou atalho
Entrada: script texto
Saida: workbook aberto, salvo, impresso ou visualizado conforme script
```

O executavel foi compilado como aplicativo Windows GUI. Isso evita abrir uma janela de CMD no uso normal. Mensagens de erro aparecem por `MessageBox`.

## 5. Estrutura do pacote

```text
sigaexcel.exe        Executavel v1
src/Program.cs       Codigo-fonte correspondente
docs/sigaexcel.md    Esta documentacao tecnica
docs/logo.png        Imagem original do icone
assets/sigaexcel.ico Icone compilado no executavel
SHA256SUMS.txt       Hashes dos arquivos do snapshot
README.md            Guia inicial do projeto
```

## 6. Uso do executavel

### 6.1 Execucao normal

```powershell
.\sigaexcel.exe C:\sigaexcel\scripts\20260617.txt
```

Executa o script no Excel usando o modo compativel padrao.

### 6.2 Ajuda

```powershell
.\sigaexcel.exe --help
```

Mostra a tela de uso.

### 6.3 Dry-run

```powershell
.\sigaexcel.exe --dry-run C:\sigaexcel\scripts\20260617.txt
```

Le e interpreta o script, mas nao abre nem automatiza o Excel. Esse modo e util para validar o parser e identificar comandos com parametros especiais.

Quando encontra `Escreve` com 13 parametros, o dry-run exibe tambem:

```text
row=<linha> col=<coluna> borda=<S/N> corBorda=<valor> corFundo=<valor>
```

### 6.4 Modo rapido

```powershell
.\sigaexcel.exe --fast C:\sigaexcel\scripts\007124.txt
```

No modo rapido, o SigaExcel executa o script com ajustes temporarios de performance:

```text
Excel.Visible = false durante a geracao
ScreenUpdating = false
EnableEvents = false
DisplayAlerts = false
Calculation = manual
Overlay de progresso visivel
Restauracao das configuracoes ao final
Excel visivel novamente no final
```

O objetivo e reduzir o tempo de geracao sem alterar o resultado funcional do relatorio.

### 6.5 Modo moderno experimental

```powershell
.\sigaexcel.exe --modern C:\sigaexcel\scripts\20260617.txt
```

Executa o script normalmente e, ao final, aplica uma formatacao visual moderna ao workbook.

Tambem e possivel salvar em um caminho especifico:

```powershell
.\sigaexcel.exe --modern --output C:\sigaexcel\planilhas\SigaExcelModern.xlsx C:\sigaexcel\scripts\20260617.txt
```

O modo moderno e experimental. Ele nao altera a interpretacao dos comandos do script, mas modifica a apresentacao visual do arquivo ao final.

### 6.6 Modo rapido com modo moderno

```powershell
.\sigaexcel.exe --fast --modern C:\sigaexcel\scripts\20260617.txt
```

Combina a execucao acelerada com a aplicacao da formatacao moderna.

### 6.7 Teste de erro

```powershell
.\sigaexcel.exe --test-error
```

Exibe uma janela de erro de teste para validar o comportamento visual de mensagens.

## 7. Modelo de execucao

O SigaExcel funciona como um interpretador sequencial:

1. Recebe o caminho do script.
2. Le o arquivo usando `Encoding.Default`.
3. Separa comandos respeitando a regra de quebra de linha.
4. Identifica o delimitador do comando.
5. Normaliza o nome do comando sem diferenciar maiusculas e minusculas.
6. Executa cada comando por automacao COM no Excel.
7. Em caso de erro, mostra arquivo, linha, comando, mensagem e texto original.

## 8. Formato dos scripts

### 8.1 Delimitadores

Os scripts podem usar `;` ou `|`:

```text
AbreExcel;S;
AbreExcel|S|
AbreArquivo|C:\SigaExcel\Modelos\Export_pc.xlsx|
Escreve;001;001;Texto;1;Arial;10;S;N;N;10;N;
```

O delimitador usado e o primeiro `;` ou `|` encontrado no comando.

### 8.2 Quebras de linha

Para preservar compatibilidade com scripts antigos, a regra da v1 e:

```text
Se o arquivo contem CR ou CRLF: separar comandos por CR e preservar LF interno.
Se o arquivo nao contem CR: separar comandos por LF.
```

Isso permite preservar conteudo de celula com `LF` dentro do campo:

```text
Escreve|002|09|'LINHA 1
LINHA 2
LINHA 3|1|Calibri|11|N|N|N|10|N|
```

### 8.3 Campos vazios

Campos vazios sao aceitos quando fazem parte do formato do comando:

```text
Escreve;006;01;;1;Arial;08;S;N;N;10;N;10;02;
```

### 8.4 Zeros a esquerda

Valores numericos com zeros a esquerda sao aceitos:

```text
001 -> 1
006 -> 6
02  -> 2
```

Isso vale para linha, coluna, sheet, tamanho e indices de cor.

### 8.5 Booleanos

Valores booleanos usam o padrao historico:

```text
S = verdadeiro
N = falso
```

Qualquer comparacao desse tipo e feita sem diferenciar maiusculas e minusculas.

### 8.6 Texto com apostrofo

O SigaExcel preserva o conteudo recebido. Se o script envia apostrofo para forcar texto no Excel, o valor e passado ao Excel dessa forma:

```text
'0001
'4106251210039
```

## 9. Comandos suportados na v1

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

Comandos conhecidos no legado, mas ainda nao implementados na v1:

```text
FormataNumeroColuna
FormataSelecao
IncluiComentario
AlinhaCelulas
```

## 10. Referencia dos comandos

### 10.1 AbreExcel

Formato:

```text
AbreExcel;
AbreExcel;S;
AbreExcel|S|
```

Parametro:

```text
1. Visivel: S/N, opcional
```

Comportamento:

- cria uma instancia de `Excel.Application`;
- define `Excel.Visible`;
- se o parametro estiver ausente, assume visivel;
- se `--fast` estiver ativo, forca Excel invisivel durante a execucao.

### 10.2 AbreArquivo

Formato:

```text
AbreArquivo|C:\SigaExcel\Modelos\Export_pc.xlsx|
```

Parametro:

```text
1. Caminho do arquivo
```

Comportamento:

```text
Excel.Workbooks.Open(caminho)
```

Observacao: caminhos absolutos legados sao preservados.

### 10.3 AdicionaSheet

Formato:

```text
AdicionaSheet;1;
AdicionaSheet;A;
```

Parametro:

```text
1. Sequencial informado pelo script
```

Comportamento da v1:

- se ainda nao houver workbook ativo, cria um workbook com uma planilha;
- se ja houver workbook ativo, adiciona uma nova planilha;
- o parametro e aceito por compatibilidade, mas nao e tratado como quantidade de abas.

Essa decisao veio de scripts reais que usam sequenciais alfanumericos:

```text
AdicionaSheet;1;
AdicionaSheet;2;
...
AdicionaSheet;A;
...
AdicionaSheet;J;
```

### 10.4 Escreve

Formatos aceitos:

```text
Escreve;linha;coluna;conteudo;
Escreve;linha;coluna;conteudo;sheet;fonte;tamanho;negrito;italico;sublinhado;corLetra;borda;
Escreve;linha;coluna;conteudo;sheet;fonte;tamanho;negrito;italico;sublinhado;corLetra;borda;corBorda;corFundo;
```

Parametros do formato completo:

```text
1. Linha
2. Coluna
3. Conteudo
4. Sheet
5. Fonte
6. Tamanho
7. Negrito
8. Italico
9. Sublinhado
10. Cor da letra
11. Borda
12. Cor da borda
13. Cor de fundo
```

Comportamento:

- localiza a celula por `Sheets(sheet).Cells(linha, coluna)`;
- define `Value` com o conteudo;
- aplica fonte, tamanho, negrito, italico, sublinhado e cor;
- se `borda=S`, aplica borda;
- se `corBorda` existir e `borda=S`, aplica borda usando `ColorIndex`;
- se `corFundo` existir, aplica preenchimento usando `Interior.ColorIndex`.

Regra importante da v1:

```text
corBorda e corFundo sao convertidos explicitamente para inteiro.
```

Exemplos:

```text
10 -> ColorIndex 10
02 -> ColorIndex 2
15 -> ColorIndex 15
20 -> ColorIndex 20
```

Valores aceitos para `ColorIndex`:

```text
1..56
-4142  xlColorIndexNone
-4105  xlColorIndexAutomatic
```

### 10.5 EscreveFormula

Formato:

```text
EscreveFormula;linha;coluna;formula;sheet;fonte;tamanho;negrito;italico;sublinhado;corLetra;borda;
EscreveFormula;linha;coluna;formula;sheet;fonte;tamanho;negrito;italico;sublinhado;corLetra;borda;corBorda;corFundo;
```

Comportamento:

- usa a mesma estrutura de `Escreve`;
- grava o conteudo em `FormulaLocal`;
- aplica as mesmas regras de fonte, borda e preenchimento.

Ponto de atencao: `FormulaLocal` depende do idioma e da localidade do Excel instalado.

### 10.6 MergeCelulas

Formato:

```text
MergeCelulas;linhaInicial;colunaInicial;linhaFinal;colunaFinal;sheet;
```

Comportamento:

```text
Range(Cells(linhaInicial, colunaInicial), Cells(linhaFinal, colunaFinal)).Merge()
```

### 10.7 FormataNumeroCelula

Formato:

```text
FormataNumeroCelula;linha;coluna;casas;decimais;sheet;
```

Comportamento:

- monta uma mascara numerica;
- aplica em `NumberFormat` da celula.

Exemplo conceitual:

```text
casas=12, decimais=2 -> ############.00
```

Esse comportamento ainda deve ser comparado com o legado para confirmar a mascara exata.

### 10.8 AutoFormata

Formato:

```text
AutoFormata;colunaInicial;colunaFinal;sheet;
```

Comportamento da v1:

- seleciona o intervalo de colunas informado;
- executa `AutoFit`.

Observacao: no binario legado, a documentacao interna sugere "tipo da auto formatacao" como terceiro parametro. Nos scripts analisados, o terceiro parametro se comporta como `sheet`. A v1 usa a interpretacao observada nos scripts reais.

### 10.9 ConfiguraPagina

Formato:

```text
ConfiguraPagina;P;1;
ConfiguraPagina;R;1;
```

Parametros:

```text
1. Orientacao: P=Paisagem, R=Retrato
2. Sheet
```

Comportamento:

```text
P -> xlLandscape
R -> xlPortrait
```

### 10.10 RenomeiaSheet

Formato:

```text
RenomeiaSheet;1;OP00694700101;
```

Parametros:

```text
1. Sheet
2. Novo nome
```

Comportamento:

```text
Sheets(sheet).Name = novoNome
```

### 10.11 GravaArquivo

Formato:

```text
GravaArquivo|C:\SigaExcel\Planilhas\Relatorio.xlsx|
```

Parametro:

```text
1. Caminho de saida
```

Comportamento:

```text
ActiveWorkbook.SaveAs(caminho)
```

Observacao: scripts legados podem informar caminhos sem extensao. A v1 preserva o caminho informado e deixa o Excel tratar o salvamento.

### 10.12 FechaExcel

Formato:

```text
FechaExcel;
```

Comportamento:

- executa `Excel.Quit`;
- libera a referencia COM principal do Excel;
- limpa as referencias internas de workbook e Excel.

### 10.13 Imprime

Formato:

```text
Imprime;
```

Comportamento:

```text
ActiveWorkbook.PrintOut()
```

### 10.14 VisualizaImpressao

Formato:

```text
VisualizaImpressao;
```

Comportamento:

```text
ActiveWorkbook.PrintPreview()
```

## 11. Tratamento de erros

Quando ocorre erro durante a execucao de um comando, a v1 monta uma mensagem com:

```text
Arquivo
Linha logica
Comando
Mensagem
Texto original do comando
Contato de suporte
```

Exemplo de estrutura:

```text
Erro ao executar script.
Arquivo: C:\sigaexcel\scripts\20260617.txt
Linha: 25
Comando: Escreve
Mensagem: Cor de Fundo fora do intervalo de ColorIndex suportado: 99
Texto: Escreve;006;01;;1;Arial;08;S;N;N;10;N;10;99;
```

No modo normal, a mensagem aparece em janela. No `--dry-run`, a saida e enviada ao console.

## 12. Modo moderno

O modo moderno aplica, ao final do script, uma formatacao geral nas planilhas do workbook ativo:

```text
Fonte Segoe UI
Tamanho base 9
Bordas cinza claras
Areas de dados com fundo branco
Faixas de cabecalho em azul
Ajustes fixos de largura de colunas
Ajustes fixos de altura de linhas
Orientacao paisagem
FitToPagesWide = 1
FitToPagesTall = 1
```

Esse modo e experimental porque pode sobrescrever estilos aplicados pelo script legado.

## 13. Modo rapido

O modo rapido altera temporariamente propriedades globais do Excel:

```text
ScreenUpdating
EnableEvents
DisplayAlerts
Calculation
Visible
```

Antes de alterar, a v1 tenta salvar o estado anterior. Ao final ou em caso de erro, tenta restaurar esse estado.

Ponto de atencao: se o processo for encerrado de forma abrupta pelo Windows ou pelo usuario, a restauracao pode nao ocorrer. Nesse caso, feche e abra o Excel novamente.

## 14. Validacoes realizadas

Validacoes registradas para a v1:

```text
Execucao real:
C:\sigaexcel\build\sigaexcel.exe C:\sigaexcel\scripts\20260617.txt
Resultado: execucao concluida sem erro
```

Verificacao manual no workbook:

```text
A6  -> Interior.ColorIndex = 15, Borders.ColorIndex = 10
G9  -> Interior.ColorIndex = 15, Borders.ColorIndex = 10
A10 -> Interior.ColorIndex = 2,  Borders.ColorIndex = 10
```

Teste do modo rapido:

```text
Comando: C:\sigaexcel\build\sigaexcel.exe --fast C:\sigaexcel\scripts\007124.txt
Resultado: execucao concluida sem erro
Tempo medido com overlay: aproximadamente 10,59 segundos
```

Teste do modo moderno:

```text
Comando: C:\sigaexcel\build\sigaexcel.exe --modern --output C:\sigaexcel\planilhas\SigaExcelModern_20260617.xlsx C:\sigaexcel\scripts\20260617.txt
Resultado: execucao concluida sem erro
```

Teste de erro visual:

```powershell
C:\sigaexcel\build\sigaexcel.exe --test-error
```

## 15. Pontos pendentes

Pontos que ainda exigem comparacao direta com o executavel legado:

- comportamento exato de `AbreExcel` sem parametro em todos os ambientes;
- parametros extras de `AbreArquivo`;
- parametros extras de `GravaArquivo`;
- montagem exata de `NumberFormat`;
- semantica original completa de `AutoFormata`;
- suporte futuro a `FormataNumeroColuna`;
- suporte futuro a `FormataSelecao`;
- suporte futuro a `IncluiComentario`;
- suporte futuro a `AlinhaCelulas`;
- comportamento de impressoes em impressoras reais;
- efeitos de idioma/localidade em `FormulaLocal`;
- comportamento com Excel 32-bit versus 64-bit;
- comportamento com arquivos protegidos, bloqueados ou em modo Protected View.

## 16. Diretrizes para evolucao

Qualquer evolucao deve preservar o modo compativel como padrao.

Recursos sugeridos para versoes futuras:

- `--validate`: validar o script sem executar Excel;
- `--log arquivo.log`: gravar log estruturado;
- `--strict`: falhar em parametros desconhecidos ou comandos incompletos;
- `--compat`: forcar comportamento legado quando houver alternativa moderna;
- suporte a caminhos relativos ao arquivo de script;
- testes automatizados por comando;
- matriz de compatibilidade por versao do Excel;
- documentacao de comandos diretamente em `sigaexcel.exe --help`.

## 17. Hash do executavel v1

Hash SHA-256 registrado para o executavel do snapshot:

```text
88CE7000560258CC1CB5A84DBEEDC3B13C2EEEB70CFA61F12AA75C6253DAE0F9
```

Comando para verificacao:

```powershell
Get-FileHash .\sigaexcel.exe -Algorithm SHA256
```
