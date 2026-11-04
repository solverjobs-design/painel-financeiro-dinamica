# Painel financeiro Dinâmica

Painel online com a análise financeira de Dinâmica Assessoria, Dinâmica Hospedagem e Dinâmica Psicologia, montado a partir da planilha de fluxo de caixa que fica no SharePoint do grupo. Ele se atualiza sozinho **todo dia às 6h (horário de Brasília)**, sem depender de nenhum computador ligado.

- **Link do painel:** https://solverjobs-design.github.io/painel-financeiro-dinamica/
- **Repositório:** https://github.com/solverjobs-design/painel-financeiro-dinamica
- **Acesso:** sem senha, por decisão da Solver em 04/11/2026: quem tiver o link vê o painel inteiro. O link não deve circular fora da Solver e do cliente. Para voltar a pedir senha, veja "Voltar a pedir senha" em Tarefas do dia a dia.

## O que o painel mostra

1. **Resultado do ano:** receitas x despesas mês a mês (realizado, mês corrente e previsto), relação despesas ÷ receitas, saldo do mês e acumulado, resultado por natureza (operação, bancos, sócios, mútuo) e uma leitura do consultor recalculada a cada atualização.
2. **Despesas:** categorias e microcategorias, custos que subiram ou caíram entre dois meses, mapa de variação, impostos, cartão de crédito (PF x PJ, cartões e estabelecimentos) e empréstimos (pagamentos por banco e contrato, tabela de contratos e pontos de atenção da dívida).
3. **Receitas:** entradas por cliente e por categoria, separando a receita da operação do dinheiro de empréstimos e sócios.
4. **Mês corrente:** a pagar x a receber por semana e as contas em aberto.
5. **Lançamentos:** tabela completa com busca, filtros e exportação em CSV.

O mês corrente é o mês da última data com lançamento **REALIZADO** na planilha (a "data-base"). Os meses anteriores contam como fechados e os seguintes como previstos. Isso avança sozinho: não há mês fixo no código.

## O que o cliente precisa fazer

Só continuar preenchendo a planilha como sempre, **no mesmo arquivo** do SharePoint. O painel lê o arquivo todo dia às 6h e mostra **"Dados atualizados em DD/MM/AAAA HH:MM"** no topo.

Para o painel ficar bom, vale manter três hábitos na planilha:
- preencher **CATEGORIA** e **SUBCATEGORIA** com os códigos do plano de contas (ex.: `3.2 Utilidades` / `3.2.01 Energia Elétrica`). Linha sem categoria válida aparece como "Não classificado";
- manter a aba **Creditos** em dia (parcelas pagas, situação, contratos novos). O painel compara essa aba com o fluxo e aponta as diferenças;
- escrever o número do contrato na descrição dos pagamentos de empréstimo sempre que o banco informar.

## Como funciona

1. Todo dia às 6h, o **GitHub Actions** roda o `gerar.py`.
2. O script baixa a planilha pelo link de compartilhamento (secret `PLANILHA_URL`) e lê três abas: **Fluxo BASE** (lançamentos), **Creditos** (contratos) e **de_para** (categoria de cada subcategoria, usada quando a categoria está em branco). As demais abas (tabelas dinâmicas, cópias, "Detalhes") não entram, para não contar nada em dobro.
3. Ele confere os dados: abas e colunas obrigatórias, quantidade mínima de lançamentos. **Se algo estiver errado, ele para e o painel do dia anterior continua no ar.** O GitHub manda um e-mail de falha para o dono do repositório.
4. Com os dados ok, ele classifica cada lançamento (natureza, categoria gerencial, conta, cartão, PF/PJ, fixo/variável, imposto, contrato de empréstimo), aplica as regras abaixo, monta o painel a partir do `modelo.html` e publica no **GitHub Pages**. Com `"senha": true` no `config.json`, o conteúdo sai criptografado com a senha (AES-256) e o link abre numa tela de senha.

O repositório guarda só o código e as regras. Os dados da planilha não ficam no repositório nem nos logs. Atenção: com o painel sem senha, os números ficam legíveis para qualquer pessoa que abrir o link.

### Regras de tratamento (todas automáticas)

| Tema | Regra |
|---|---|
| Escopo | Lançamentos do ano da data-base. Saldo inicial, linhas sem data ou valor e empresas fora de `"empresas"` (config.json) ficam de fora e são contados nas notas do painel. |
| Categoria gerencial | Pelo código do plano de contas (`"plano_de_contas"`). Vale o código da CATEGORIA; o da SUBCATEGORIA só refina quando é filho dela (ex.: 11.1.03 Aplicações). Textos como RETIRADA e DL têm regra própria (`"regras_por_texto"`). |
| Mútuo | 12.6 com empresa do grupo (nome contém `"nomes_do_grupo"`) = mútuo entre empresas; com qualquer outra pessoa = mútuo com sócios. |
| Transferências | Saída ou entrada sem categoria, por PIX/TED, que bate em data e valor com o lado oposto em outra conta vira movimentação interna. |
| Cartão | Nos meses em que as compras de um cartão estão detalhadas, o pagamento da fatura desse cartão é retirado (sem dupla contagem). Sem detalhe, a fatura fica como "Cartão de crédito (fatura)". |
| Empréstimos | Ligados aos contratos da aba Creditos pelo número do contrato; sem número, pelo valor da parcela (mesma empresa e banco) ou pelo nome. Previstos sem número usam o valor da última parcela paga (até 3% de diferença). Separados em parcela, custos de contratação, liquidação antecipada e liberação. |
| Pontos de atenção | Calculados a cada dia: parcelas pagas na aba x fluxo, contratos marcados como quitados com parcelas previstas, pagamentos sem contrato na aba, previsões em duplicidade e aumento de parcela após carência. |

## Tarefas do dia a dia

| Quero... | Como fazer |
|---|---|
| Atualizar agora | GitHub → aba **Actions** → **Atualizar painel** → **Run workflow** |
| Ver se rodou | GitHub → aba **Actions** (verde = ok, vermelho = falhou) |
| Pausar a atualização | Actions → Atualizar painel → menu **⋯** → **Disable workflow** |
| Retomar | Actions → Atualizar painel → **Enable workflow** |
| Voltar a pedir senha | Em `config.json`, troque `"senha": false` por `"senha": true`. A senha usada é a do secret `PAINEL_SENHA` (já cadastrado). A mudança vale assim que o commit é enviado. |
| Trocar a senha (com senha ligada) | Settings → Secrets and variables → Actions → `PAINEL_SENHA` → **Update**. Depois rode **Run workflow**. Quem marcou "lembrar" terá que digitar a senha nova. |
| Trocar o arquivo da planilha | Gere um novo link "qualquer pessoa com o link" no SharePoint e atualize o secret `PLANILHA_URL`. Depois rode **Run workflow**. |
| Mudar o horário | Em `.github/workflows/atualizar.yml`, altere `cron: "0 9 * * *"`. O horário é UTC: Brasília + 3 h (6h → 9, 7h → 10). |
| Renomear uma aba | Em `config.json`, ajuste `"abas"`. |
| Incluir uma empresa nova do grupo | Em `config.json`, acrescente o nome (como aparece na coluna EMPRESA) em `"empresas"`. |
| Ajustar classificação | Em `config.json`: `"plano_de_contas"` (código → categoria e natureza), `"custos_fixos"`, `"impostos"`, `"bancos"`, `"apelidos"`. |
| Fixar o ano | Em `config.json`, troque `"ano": "auto"` por `"ano": 2026`. Com "auto", vale o ano da data-base. |

## Configuração (feita uma vez)

**Secrets do repositório** (Settings → Secrets and variables → Actions → New repository secret):

- `PLANILHA_URL`: link de compartilhamento da planilha no SharePoint ("qualquer pessoa com o link pode exibir").
- `PAINEL_SENHA`: senha do painel (mínimo de 8 caracteres; de preferência 4 palavras aleatórias). Só é usada com `"senha": true` no `config.json`.

**Planilha:** compartilhada no SharePoint como "qualquer pessoa com o link". Quem tiver esse link consegue baixar a planilha, por isso ele fica guardado só no secret e não deve circular. Se a TI da Dinâmica desligar esse tipo de link, o caminho é um aplicativo no Microsoft Entra ID com permissão de leitura (exige um administrador da Microsoft 365 do cliente).

**GitHub Pages:** Settings → Pages → Source: **GitHub Actions**.

## Mensagens de erro comuns (aba Actions)

- `o secret PLANILHA_URL não está cadastrado` / `o secret PAINEL_SENHA não está cadastrado` → falta o secret.
- `o link da planilha não devolveu um arquivo Excel` → o compartilhamento "qualquer pessoa com o link" foi desligado ou o link mudou.
- `aba "..." não encontrada` → a aba foi renomeada; ajuste `config.json`.
- `a coluna "..." sumiu da aba` → um cabeçalho foi renomeado ou apagado.
- `a aba de fluxo tem só N lançamentos` → a aba foi esvaziada ou filtrada; confira a planilha.

## Teste local (opcional)

```bash
pip install -r requirements.txt
python gerar.py --xlsx copia-da-planilha.xlsx
```

O resultado fica em `site/index.html`. Com a senha ligada, rode com `PAINEL_SENHA=uma-senha-de-teste` e abra por um servidor local (por exemplo, `python -m http.server --directory site`), porque a descriptografia exige `https` ou `localhost`. Com `--json dados.json` o script grava também os dados tratados, sem criptografia, para conferência (não suba esse arquivo).

## Arquivos

| Arquivo | Para quê |
|---|---|
| `config.json` | senha ligada ou não, abas, empresas, plano de contas e demais regras de classificação |
| `gerar.py` | baixa a planilha, valida, trata e gera o painel (criptografado quando a senha está ligada) |
| `modelo.html` | layout, gráficos, tela de senha e textos do painel |
| `.github/workflows/atualizar.yml` | agendamento diário e publicação |
