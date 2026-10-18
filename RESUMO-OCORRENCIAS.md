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
python -m unittest discover -s .  # 161 testes
```

No Linux o Tkinter vem separado: `sudo apt install python3-tk`. No Windows já vem
com o instalador do python.org.

Para gerar o `.exe`:

```bash
python -m PyInstaller --onefile --windowed --name "Resumo de Ocorrencias" \
    --add-data "admin.json:." excel_filter.py
```

Sem o `admin.json` ao lado do script, o `.exe` sai sem admin: o login responde
que não há admin configurado.

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

O app inteiro é escuro. No topo, o botão **Unilever** e, para o admin, Abrir,
+ Adicionar, API e o histórico; embaixo, um painel com os filtros numa linha só — datas, turno,
linhas, palavra-chave — e Filtrar / Baixar Dados à direita.
Abaixo dos filtros, o título **Oráculo Digital** com o subtítulo "enxergando
oportunidades" e, sob eles, o período filtrado.

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
- **Ocorrências por equipamento**, logo abaixo dos tiles: cada linha do relato
  começa com a máquina ("Gasser - alarme…") e vira uma contagem. Passar o mouse destaca;
  **clicar leva para as ocorrências daquele equipamento**, já filtradas.
- **Tempo de troca (changeover) e produção por dia**, num gráfico só: em cima,
  as barras de minutos de changeover; embaixo, numa faixa, a produção do dia —
  azul bateu o plano, laranja ficou abaixo, ardósia sem plano anotado — com a
  marca branca no plano. Acima de 31 dias o eixo vira mês. **Clicar num dia** o
  destaca e leva o painel do lado até ele; **duplo clique** filtra só aquele dia.
- **SKU e changeover por dia**, lado a lado com o gráfico: o SKU inicial do dia
  e, a cada troca, "houve um changeover 08h - 09h (31m19s) · passa a 64830874".
  Com três trocas, chega ao terceiro SKU.

Passar o mouse num dia mostra changeover, produzido, plano e % exatos. A janela
abre maximizada.

### Unilever, Deslogar e Adicionar API (acesso e base salva)

Quem abre o app é **leitor**: vê o botão **Unilever**, os filtros, as abas e o
Baixar Dados, com a **base salva** já carregada. Unilever abre o login; com a
conta de admin, o mesmo botão vira **Deslogar** e aparecem **Abrir**,
**+ Adicionar**, o histórico e, ao lado do Deslogar, **Adicionar API**. O login
não fica lembrado: cada vez que o app abre, começa como leitor.

**Adicionar API** tem duas abas:

- **Fontes em tempo real** — **Enviar .zip...** importa uma vez. O **token /
  código** é o token de acesso do Microsoft Graph (o do Graph Explorer serve).
  Cada **fonte** é um nome (ex.: outubro, 2026, Linha 1) e uma **chave**:
  `drives/<id>/items/<id>` de um **arquivo ou de uma pasta** — a do ano ou a do
  mês; o app percorre ANO → MÊS → ARQUIVO, pega só `.xlsx`/`.zip` (ignora o `~$`
  de trava) e tira o mês dos nomes das pastas ("2026" / "10 - Outubro") —, a URL
  completa do Graph com ela (`/content` e `/children` também servem), o link do
  SharePoint, ou o arquivo/pasta que o OneDrive sincroniza. **Testar conexão** só
  lê: diz se o token e a chave abrem o arquivo ou a pasta e quais planilhas e
  meses estão lá, sem baixar nem salvar nada. Enquanto o app está aberto, ele confere cada fonte **a cada minuto**
  e só baixa quando o arquivo mudou (o `cTag` do Graph, o ETag do link, ou a data
  e o tamanho do arquivo sincronizado). Mudou — "Ricardo trocou x por z" — a base
  é atualizada e a tela se refaz **mantendo os filtros** de quem está olhando, com
  "Atualizado às 14:32: outubro (modificado por Ricardo)" na barra de status.
- **Na base** — uma linha por planilha (linha de produção, mês, linhas, origem,
  quando entrou); marque várias e **Apagar selecionados**, com confirmação.

A mesma linha no mesmo mês **substitui** a anterior: reenviar o zip, mesmo salvo
como "(1)", ou o mesmo mês chegar de novo pela fonte, não dobra nenhum total.

Tudo fica em `~/.resumo-ocorrencias/base/` (no Windows,
`C:\Users\<usuário>\.resumo-ocorrencias\base`): um índice e um arquivo por
planilha, já lidos — abrir o app não relê os zips — e `fontes.json` com as fontes
e o token. Fechar o app não apaga nada; só o admin apaga.

O app **só lê** do SharePoint — todo pedido é GET: descrever o arquivo, listar a
pasta, baixar. Nada é alterado nem apagado lá; Apagar e Remover fonte mexem só
na cópia do app. Pelo Graph, o token vai só nos pedidos ao Graph; o download usa
o endereço pré-assinado que o Graph devolve, sem token (como a documentação
manda), e a lista de uma pasta só segue a próxima página se ela for no Graph.
Um token vencido aparece como "O token expirou" na fonte, uma vez, e as outras
fontes continuam.

A senha nunca é gravada: `admin.json` guarda o login, um sal e o hash PBKDF2. O
arquivo fica **fora do git** (o repositório é público) e entra no `.exe` na hora
do build. Para criar ou trocar: `python excel_filter.py --admin <login>` pede a
senha duas vezes e grava o `admin.json` ao lado do script.

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
- `Seladora` e `Seladora de Caixas` contam separado, porque é assim que está
  escrito no relatório. Unificar exigiria um dicionário de apelidos de máquina.
- O plano é lido como está escrito; o app **não** recalcula a partir da tabela de
  formatos (a aba Planejamento tem três blocos com valores diferentes para o
  mesmo formato e nada diz qual vale quando).
- O login de admin é uma porta na tela, não segurança: quem desempacotar o
  `.exe` passa por ela, e a base é JSON comum no disco. Proteção de verdade pede
  um servidor decidindo quem escreve.
- A base é **deste computador**. Em outro PC o mesmo `.exe` começa com a base
  vazia.
- O token do Graph colado à mão vale cerca de **1 hora**; depois a fonte para com
  "O token expirou". Tempo real por meses sem ninguém colar token exige registrar
  o app no Azure AD da empresa (TI), para ele renovar o acesso sozinho — ou usar
  como chave o arquivo de uma pasta que o OneDrive sincroniza, que não precisa de
  token nenhum.
- "Tempo real" é conferir a cada minuto com o app aberto: um app de desktop não
  recebe aviso do SharePoint (webhook pede endereço público).
- O token fica em texto puro em `fontes.json`, na pasta do usuário.
- O comportamento do Graph foi testado contra um servidor local escrito a partir
  da documentação da Microsoft, não contra o SharePoint da Unilever.

---

## 8. Testes

161 testes com `unittest`, sem framework extra. Cobrem: leitura das abas por dia
e das versões do modelo, mês de referência e conflitos, células mescladas, filtro
por palavra/período/turno/linha, junção de pacotes, histórico, totais e
desconhecidos, tempo de troca, sequência de SKU, gráficos (contando os retângulos
desenhados num Tk de verdade), chips, máscara de data, a equivalência entre o
leitor rápido e o lento, o login de admin, a base salva (ida e volta, reenvio sem
duplicar, apagar, índice adulterado) e as fontes em tempo real — link com ETag e
304, Graph com o token só nos metadados e o `u!` do exemplo da documentação,
arquivo sincronizado, fonte quebrada que não para as outras — contra servidores
HTTP locais de verdade.
