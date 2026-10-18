# Resumo de Ocorrências — documentação completa

App em Python (Tkinter + openpyxl) que lê os **Relatórios de Turno** em Excel e
responde, com filtros, o que aconteceu na linha de produção: ocorrências por
período, palavra-chave, turno e linha; e a produção por trás delas.

Entregue como **`.exe` único** para Windows (12 MB, sem instalação, sem Python na
máquina de quem usa). Código em inglês, tela em português.

---

## 1. Como rodar

```bash
pip install -r requirements.txt
python excel_filter.py            # abre a tela
python -m unittest discover -s .  # 125 testes
```

No Linux o Tkinter vem separado: `sudo apt install python3-tk`. No Windows já vem
com o instalador do python.org.

Para gerar o `.exe`:

```bash
python -m PyInstaller --onefile --windowed --name "Resumo de Ocorrencias" excel_filter.py
```

O título da janela carrega o carimbo do build (`build 15/09 01:18`) — é como se
sabe qual versão está rodando.

---

## 2. O que entra

- **Um `.xlsx`** — um mês de relatório de turno.
- **Um `.zip`** com vários `.xlsx` — vários meses de uma vez.
- **Vários pacotes somados** — abrir um segundo arquivo pergunta "somar ou
  substituir?"; somando, D01, D02 e D03 convivem na mesma tela e a coluna Linha
  permite compará-los ou isolar um.

Medido nos pacotes reais: D01 (8 planilhas, 527 MB) abre em **19 s**; os três
pacotes somados dão **24 planilhas, 17.493 linhas, 16.001 ocorrências**,
01/2026 a 08/2026.

Toda planilha aberta entra num **histórico** (`~/.resumo-ocorrencias/historico.json`,
as 15 últimas) com data, linhas e mês. O **✕** ao lado de Reabrir tira o item
escolhido; sem item escolhido, limpa tudo.

---

## 3. A tela

O app inteiro é escuro. No topo, a barra do arquivo (Abrir, + Adicionar,
histórico); embaixo, um painel com os filtros numa linha só — datas, turno,
linhas, palavra-chave — e Filtrar / Baixar Dados à direita.

O app abre na aba **Produção**; Ocorrências é a segunda. Acima das abas fica só
o período filtrado ("05/01/2026 a 31/08/2026").

### Aba Ocorrências

Filtros: **data inicial / data final**, **turno** (1º 06h–14h, 2º 14h–22h, 3º
22h–06h, pela hora em que a linha começa), **linhas** (chips com ✕) e
**palavra-chave**. Todos se somam.

- Palavra-chave é "contém", sem diferenciar maiúsculas; separando por vírgula,
  basta **uma** casar. Com palavra-chave preenchida, aparece **só a frase que a
  contém**, não o relato inteiro da hora — e a palavra fica marcada em amarelo.
- A lista sai **agrupada por linha de produção**, um bloco por linha.
- Data aceita `06/11/2023`, `06112023`, `06-11-2023`, `2023-11-06`, `06/11/23` e
  **só o dia** (`6`), que vale como aquele dia do mês do relatório. A barra é
  inserida enquanto se digita.

### Aba Produção

Os mesmos filtros, vistos como números:

- **Tiles**: planejado, produzido, faltou (laranja quando falta), % atingido,
  linhas de ocorrência.
- **Tempo de troca (changeover) e produção por dia**, num gráfico só: em cima,
  as barras de minutos de changeover; embaixo, numa faixa, a produção do dia —
  azul bateu o plano, laranja ficou abaixo, ardósia sem plano anotado — com a
  marca branca no plano. Acima de 31 dias o eixo vira mês. **Clicar num dia** o
  destaca e leva o painel do lado até ele; **duplo clique** filtra só aquele dia.
- **SKU e changeover por dia**, lado a lado com o gráfico: o SKU inicial do dia
  e, a cada troca, "houve um changeover 08h - 09h (31m19s) · passa a 64830874".
  Com três trocas, chega ao terceiro SKU.
- **Ocorrências por equipamento**, embaixo: cada linha do relato começa com a
  máquina ("Gasser - alarme…") e vira uma contagem. Passar o mouse destaca;
  **clicar leva para as ocorrências daquele equipamento**, já filtradas.

Passar o mouse num dia mostra changeover, produzido, plano e % exatos. A janela
abre maximizada.

### Baixar Dados

Botão embaixo de Filtrar: grava um `.xlsx` novo com **Linha, Data, Hora e
Ocorrência** do que a aba Ocorrências mostra, cabeçalho congelado e filtro do
Excel ligado. Com palavra-chave, só a frase que a contém.

---

## 4. Como a planilha é lida (e o que ela esconde)

### As datas da "Tabela de Dados" estão erradas

A coluna `Data` foi arrastada somando **um dia por linha** em vez de por dia:
anda de 01/09 a 24/10 para um mês só. Pior, a fórmula que traz o texto escolhe a
aba do dia por `TEXT($A,"d")` — então o bloco "Dia 1" lê 24 abas diferentes, uma
por hora. **O app não usa essa aba.**

Ele lê as **abas por dia** (`1`..`31`): cada uma tem três turnos de 8 horas lado a
lado. A data é remontada como "dia = nome da aba" + mês de referência. Resultado:
720 linhas e 578 ocorrências contra 518 da tabela quebrada, que perdia 60.

### De onde vem o mês

As abas por dia não dizem o mês. O app procura uma aba que seja **calendário** —
cada dia uma vez, em sequência — e acha a `Plan.Prod`. No pacote, também vale o
**nome do arquivo** ("… - Março.xlsx"); quando os dois discordam, **o nome ganha**
e a tela mostra o conflito. Sem ano no nome, herda o ano dos outros arquivos do
pacote. Se ninguém disser o mês, a tela avisa em vez de fingir que leu.

### Ocorrência mesclada

A célula de Principais Perdas costuma ser **mesclada** sobre várias horas: um
relato que vale das 06h às 10h ocupa quatro linhas. O app mostra o período real
("06h - 10h"); para o filtro de turno vale a hora em que ele começa.

### Versões do modelo

V6, V7 e V8 põem o cabeçalho em linhas e colunas diferentes (linha 26 com blocos
B/AI/BP; linha 27 com B/AJ/BR). Os blocos são localizados pelos **próprios
rótulos** — com coordenada fixa, um mês inteiro era descartado em silêncio.

### Linha de produção

Vem do título de cada bloco ("LINHA D02 - 1º TURNO"). Quando só o 1º turno tem
título, os outros **herdam** a linha do mesmo dia — sem isso, os turnos 2 e 3 iam
para um bloco solto no fim da lista e parecia que o dia acabava às 14h.

### Formato

`Regular`, `Golias`, `Hagrid`, `Axe`, `Tripack` são **dado transportado**, não
regra: o app lê o formato e o plano escritos em cada hora. Por isso a linha azul
varia de hora em hora (Hagrid 1.482/h, Regular 1.925/h, Golias 1.314/h) e
qualquer formato novo entra sozinho.

### Changeover

A **hora** só existe no relato do turno ("Change over: Troca de formato do…") e a
**duração** só no cabeçalho do bloco (`TROCA 1 - TEMPO = 31m19s`). O app cruza os
dois. Conta como evento a linha que **começa** com a expressão — "atrasando o
change over" no meio de uma frase é contexto, não outra troca —, e aceita as
grafias que aparecem: `change over`, `changeover`, `CHANGER OVER`, `Chang Over`,
`troca de formato/variante`. Troca cronometrada sem relato aparece **sem hora**,
dizendo isso.

---

## 5. Desconhecido nunca vira zero

Regra que atravessa o projeto inteiro:

- Célula de quantidade ilegível fica **fora da soma** e o total ganha `+`.
- Hora sem plano anotado vira **barra cinza sem linha azul**, não barra verde.
- Data ilegível **interrompe** dizendo a linha, em vez de sumir calada.
- Mês adivinhado **se declara** na tela.
- Duração que não é duração vira desconhecido, não zero.

Um caso real: 113 células de "Produzido" guardam o **texto** `"0"` em vez do
número, e cinco horas de 03/02/2026 têm a célula Plan vazia.

---

## 6. Desempenho

| | antes | agora |
|---|---|---|
| um arquivo de 58 MB | 26,2 s | **1,3 s** |
| zip com 8 planilhas (527 MB) | ~4 min | **19 s** |

Dos 26 s originais, 25,6 s eram o `load_workbook`: o modo normal do openpyxl
materializa toda célula de toda aba. Em modo streaming o mesmo arquivo abre em
0,1 s, mas perde acesso aleatório e as células mescladas — então as ~70 primeiras
linhas de cada aba viram uma grade em memória e as mesclagens saem do XML da aba.
Um teste lê a mesma planilha pelos **dois caminhos** e exige linhas idênticas.

A leitura roda **fora da thread da tela**, com progresso ("Lendo 4/8: …") e os
campos travados enquanto lê. A janela não congela mais.

---

## 7. Limites conhecidos

- Só `.xlsx`; `.xls` e macros de `.xlsm` não passam pelo openpyxl.
- Mais de 2 palavras-chave: as linhas certas somem, mas o Excel só aceita 2
  critérios por coluna e o dropdown do arquivo exportado abre sem o critério.
- `Seladora` e `Seladora de Caixas` contam separado, porque é assim que está
  escrito no relatório. Unificar exigiria um dicionário de apelidos de máquina.
- O plano é lido como está escrito; o app **não** recalcula a partir da tabela de
  formatos (a aba Planejamento tem três blocos com valores diferentes para o
  mesmo formato e nada diz qual vale quando).
- O visual é o Tkinter cru, ainda sem o acabamento de cartões/paleta.

---

## 8. Testes

125 testes com `unittest`, sem framework extra. Cobrem: leitura das abas por dia
e das versões do modelo, mês de referência e conflitos, células mescladas, filtro
por palavra/período/hora/linha, junção de pacotes, histórico, totais e
desconhecidos, tempo de troca, sequência de SKU, gráficos (contando os retângulos
desenhados num Tk de verdade), chips, máscara de data, e a equivalência entre o
leitor rápido e o lento.
