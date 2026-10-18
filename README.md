# Resumo de ocorrências

App em Python (Tkinter) que lê a planilha de turno e mostra as ocorrências, com
três filtros: **período** (data inicial / data final), **palavra-chave** e
**turno** (1º 06h–14h, 2º 14h–22h, 3º 22h–06h). **Baixar Dados** grava um
`.xlsx` com Linha, Data, Hora e Ocorrência do que está filtrado.

```bash
pip install -r requirements.txt
python excel_filter.py           # abre a tela
python -m unittest discover -s . # testes
```

No Linux o Tkinter vem separado: `sudo apt install python3-tk`. No Windows já
vem com o instalador do python.org.

## Vários meses de uma vez (.zip)

Abrindo um `.zip`, o app lê **todas as planilhas de dentro** e trata como um
conjunto só. Cada arquivo diz o próprio mês: primeiro pela aba-calendário
(`Plan.Prod`), senão pelo nome do arquivo ("… - Março.xlsx"). Quando os dois
discordam, **o nome ganha** e a tela mostra o conflito — no pacote real, o
arquivo de Fevereiro tem uma `Plan.Prod` esquecida do template dizendo 11/2023.
Arquivo cujo nome traz o mês mas não o ano herda o ano dos outros do pacote.

Não existe campo de "mês do relatório": o período (data inicial/final) faz esse
papel. Ler um pacote de 6 planilhas de ~60MB leva cerca de 3 minutos — a janela
mostra qual arquivo está sendo lido.

O modelo tem versões diferentes (V6, V7, V8) com o cabeçalho em linhas e colunas
distintas, então os blocos de turno são localizados pelos próprios rótulos
("Horário", "Plan", "Produzido"…), nunca por coordenada fixa.

## Vários pacotes juntos

**+ Adicionar** soma outro arquivo ou zip ao que já está aberto: os dois viram uma
base só, e como a linha de produção é coluna, dá para comparar D01, D02 e D03 na
mesma tela. O título mostra `D01.zip + D02.zip` e o histórico guarda cada um
separado. O **✕** ao lado de Reabrir tira do histórico o item escolhido — sem
nada escolhido, limpa o histórico inteiro.

Medido no pacote real: D01 com 8 planilhas e 527MB abre em **19 segundos**, e a
janela continua respondendo durante a leitura, que roda fora da thread da tela.

A leitura usa o modo streaming do openpyxl: abrir um arquivo de 58MB em modo
normal leva 25s, em streaming leva 0,1s. Como o modo streaming não dá acesso
aleatório nem as células mescladas, as linhas do cabeçalho vão para uma grade em
memória e as mesclagens saem direto do XML da aba. Um teste compara os dois
leitores célula a célula, e nos arquivos reais o resultado é idêntico — 20x mais
rápido.

## Períodos longos

O gráfico muda de granularidade sozinho, senão vira cerca de arame: um dia por
hora; até 120 barras, por hora ao longo dos dias; até 31 dias, por dia; acima
disso, **um total por mês**.

## A data do relatório (leia isto)

A aba "Tabela de Dados" do Relatório de Turno **tem as datas erradas**. A coluna
`Data` foi arrastada somando um dia por **linha** em vez de por dia, então ela
anda de 01/09 a 24/10 para um mês só; e a fórmula que traz o texto escolhe a aba
do dia por `TEXT($A,"d")`, de modo que o bloco "Dia 1" lê 24 abas diferentes, uma
por hora. Filtrar por período em cima dela devolve o dia errado.

Por isso o app **não usa essa aba**: lê as 31 abas por dia (`1`..`31`), onde cada
uma tem os três turnos de 8 horas lado a lado (colunas B, AI e BP, cabeçalho na
linha 26). A data é remontada como "dia = nome da aba" + o mês de referência.

O mês não está nas abas por dia. O app procura uma aba que seja calendário — cada
dia uma vez, em sequência — e acha `Plan.Prod` (01/11 a 30/11/2023). A aba
quebrada diz setembro; por isso o critério é calendário, não "a primeira data que
aparecer". Se nenhuma aba parecer calendário, o app avisa na tela em vez de fingir que leu.

Resultado prático: 720 linhas (30 dias x 24h) e **578 ocorrências**, contra as
518 da tabela quebrada — que também perdia 60 ocorrências.

## O que o app descobre sozinho

Ao abrir o arquivo, ele escolhe a aba que é tabela de verdade (cabeçalho na
linha 1 e mais linhas — "Tabela de Dados" entre as 35 do relatório de turno) e
acha as colunas de **data**, **hora** e **ocorrência** pelo conteúdo, não pelo
nome: "Hora" (1..24) perde para "Horário" ("07h - 08h"), e a ocorrência é a
coluna de texto mais longa. Não há nada pra configurar.

Com palavra-chave preenchida, a lista mostra **só a frase que contém a palavra**,
não o relato inteiro da hora — a célula do Excel guarda o turno todo. Sem
palavra-chave, o texto aparece inteiro.

Toda planilha aberta entra num **histórico** (`~/.resumo-ocorrencias/historico.json`,
as 15 últimas) com data de abertura, número de linhas e mês; o seletor no topo
reabre qualquer uma com o mesmo mês de antes.

A lista mostra só as linhas que têm ocorrência escrita. A linha de resumo traz a
contagem, o período encontrado e a soma das colunas de quantidade — dia, turno e
hora são código e não entram na soma.

## Aba Produção

A segunda aba lê o mesmo filtro da primeira e mostra a produção:

- **Tiles**: planejado, produzido, faltou, % atingido e total de ocorrências.
- **Planejado x produzido por dia**: barra cinza é o plano, verde é o real.
- **Ocorrências por equipamento**: cada linha do relato começa com a máquina
  ("Gasser - alarme...") e vira uma contagem — é o "quantas vezes apareceu
  ENCAIXOTADORA" do período.
- **SKU do período**: agrupado pelo código de 8 dígitos, porque a mesma
  referência aparece escrita de três jeitos na planilha, com as horas de cada um.

Diferença importante entre as abas: a lista de ocorrências mostra só as horas com
texto escrito; a aba Produção conta **todas** as horas do filtro, senão o plano do
turno sumiria junto com as horas sem relato.

Os gráficos são desenhados em `tk.Canvas`, sem matplotlib — dependência que
triplicaria o `.exe` para desenhar retângulos.

## Regras dos filtros

Palavra-chave é "contém", sem diferenciar maiúsculas/minúsculas; separando por
vírgula, basta **uma** casar (OU). Período é inclusivo nas duas pontas e
qualquer ponta pode ficar vazia. Os três filtros se somam (E). Datas em
`dd/mm/aaaa`, `dd/mm/aa`, `aaaa-mm-dd` ou já formatadas como data no Excel.
Linha com data ilegível interrompe dizendo o número da linha, em vez de sumir
calada. Os valores lidos são os **calculados** — a coluna pode ser fórmula, como
é o caso no relatório de turno.

## Baixar Dados

Um `.xlsx` novo com quatro colunas — Linha, Data, Hora e Ocorrência — e uma
linha por ocorrência que a aba Ocorrências está mostrando. Com palavra-chave, a
ocorrência vem só com a frase que a contém. Coluna que a planilha não tem sai em
branco, nunca inventada. A planilha original nunca é aberta para escrita.

O turno é lido da hora em que cada linha **começa**: "13h - 14h" é 1º turno.

## Limites

- Só `.xlsx`; `.xls` e macros de `.xlsm` não passam pelo openpyxl.
- Ler o arquivo de 59MB leva ~20s (uma vez só, os filtros depois são instantâneos).
