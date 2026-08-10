# Dashboard Reforma Tributária — Grupo Delga

Lê o arquivo `.mpp` do MS Project diretamente (sem conversão manual), calcula os
KPIs (progresso total, SPI, marco crítico), mostra o status por etapa e permite
editar o % concluído na tela, gerando um `.xml` atualizado que abre nativamente
no MS Project (Arquivo → Abrir).

## Como funciona

- A leitura do `.mpp` é feita via **MPXJ** (biblioteca Java, usada através do
  pacote Python `mpxj` + `jpype1`) — por isso o `packages.txt` instala o Java
  (`default-jdk`) no ambiente do Streamlit Cloud.
- Etapas em que todas as tarefas ainda estão na mesma data (ex: RH, sem
  cronograma lançado) são detectadas automaticamente e excluídas dos KPIs,
  aparecendo como "aguardando cronograma" — não é uma lista fixa, se ajusta
  sozinho conforme você for preenchendo as datas reais no Project.
- A exportação **não grava em `.mpp`** (a Microsoft não documenta esse formato
  binário o suficiente para escrita externa confiável — nem o MPXJ suporta).
  Em vez disso, gera um `.xml` no formato MSPDI, que é o formato de intercâmbio
  oficial da Microsoft — o Project abre esse `.xml` normalmente como se fosse
  um arquivo nativo (Arquivo → Abrir → selecionar o `.xml`).

## Deploy (mesmo fluxo que você já usa na Plataforma Delga)

1. Suba estes arquivos (`app.py`, `requirements.txt`, `packages.txt`) para um
   repositório no GitHub (pode usar o editor do navegador, sem precisar de
   terminal).
2. Em [share.streamlit.io](https://share.streamlit.io), crie um novo app
   apontando para esse repositório e para `app.py`.
3. O Streamlit Cloud vai instalar o Java (via `packages.txt`) e as dependências
   Python (via `requirements.txt`) automaticamente no primeiro deploy — isso
   pode levar alguns minutos a mais que os seus outros apps (por causa do
   Java), mas só na primeira vez.
4. Depois de publicado, é só abrir o link, subir o `.mpp` na barra lateral e o
   dashboard é montado na hora.

## Uso

1. Abra o app → envie o `.mpp` (ou um `.xml` já exportado do Project) na
   barra lateral.
2. Confira os KPIs, o status por etapa e as atividades relevantes.
3. Se quiser atualizar o progresso, edite a coluna **% concluído** na tabela
   no final da página e clique em **"Gerar arquivo atualizado (.xml)"**.
4. Baixe o `.xml` e abra direto no MS Project.

## Limitações conhecidas (primeira versão)

- Só edita **% concluído** por enquanto — editar datas/durações direto no
  dashboard fica para uma próxima versão, se fizer sentido.
- Cada sessão do navegador precisa reenviar o `.mpp` — o app não fica
  "lembrando" o último arquivo entre visitas (dá pra evoluir isso depois,
  guardando o último upload no repositório, no mesmo padrão do Bobinas).
- O cálculo de "% previsto" usa a data de hoje contra Início/Término de cada
  tarefa — não depende de você ter clicado em "Definir Linha de Base" no
  Project (mais simples, mas menos preciso que uma baseline formal).
