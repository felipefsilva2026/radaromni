# Pesquisa com curadoria: Omni

Versão da pesquisa de mercado com **revisão humana antes de chegar à diretoria**. O projeto `PESQUISA DE MERCADO` (jornal diário automático) continua funcionando igual; este aqui é independente e não lê nem altera nada de lá.

## Fluxo

1. **`1 - Gerar planilha.bat`**: pesquisa as notícias (mesmas 39 buscas e mesmo filtro de confiabilidade de fontes do jornal) e cria `planilhas/curadoria_AAAA-MM-DD_HHMM.xlsx`, que abre sozinha no Excel.
2. **Revisão na planilha**: a equipe lê as notícias, ajusta as colunas editáveis (cabeçalho laranja) e marca **SIM** ou **NÃO** na última coluna, **STATUS**. Pode renomear o arquivo e mandar para outras pessoas filtrarem.
3. **`2 - Gerar HTML diretoria.bat`**: lê a planilha salva e gera `html_diretoria/radar_diretoria_AAAA-MM-DD_HHMM.html`, que abre sozinho no navegador. Por padrão pega a planilha mais recente da pasta `planilhas`; para usar outra (por exemplo, a que voltou do pessoal por e-mail), **arraste o arquivo em cima do `.bat`**.

## Colunas da planilha

**Todas as colunas são editáveis**, e o HTML usa o que estiver escrito:

| Coluna | Aparece no HTML? | O que é |
|---|---|---|
| Data, Título, Fonte, Link | sim | Dados da notícia. Pode corrigir título, fonte, data ou trocar o link |
| Categoria | sim | Tema da notícia. Lista suspensa, mas aceita digitar um tema novo |
| Imagem | sim | Endereço da foto. Cole outro endereço para trocar a foto, ou apague para usar o painel com ícone |
| **Impacto** | sim | Alto / Médio / Baixo, com sugestão automática |
| **Impacto para a Omni** | sim | Texto que a diretoria lê. Vem com sugestão por regra fixa (categoria + termos encontrados) |
| **Área responsável** | sim | Quem deve acompanhar o tema. Vem com sugestão por categoria |
| **Comentário da equipe** | sim | Opcional. Se preenchido, aparece em destaque no HTML |
| Score | ordena | Define a ordem dentro de cada nível de impacto (maior primeiro) |
| ID, Resumo, Nota da fonte, Prioridade sugerida | não | Referência para a revisão |
| **TOP 6** | escolhe | Posição (1 a 6) na página executiva. Vazio = não entra no Top 6 |
| **STATUS** | filtra | SIM = vai para a diretoria. NÃO ou vazio = não vai |

O HTML **sempre usa o que estiver preenchido na planilha**: as sugestões são só ponto de partida. A ordem das colunas pode mudar, mas os **nomes dos cabeçalhos não**. Um valor de impacto fora da lista aparece como "Não classificado".

## O HTML da diretoria

Uma página só, sem depender de internet (menos as fotos das notícias), com:

- números principais (notícias relevantes, alto impacto, Omni na mídia, áreas envolvidas);
- visão geral: barra por nível de impacto, notícias por tema e áreas que devem acompanhar;
- **Destaques**: as notícias de alto impacto em cards grandes, com "o que significa para a Omni" e o comentário da equipe;
- todas as notícias selecionadas, por tema, da mais para a menos impactante.

Dá para imprimir ou salvar como PDF pelo navegador (Ctrl+P).

### Fotos das notícias

O RSS do Google não traz foto. Por isso, no passo 1, cada link do Google é convertido no endereço real da matéria (que vai para a coluna **Link** da planilha), e é lida a foto de destaque que o próprio site publica para compartilhamento (a mesma que aparece ao mandar o link no WhatsApp). Isso roda em paralelo e acrescenta alguns minutos ao passo 1.

No passo 2, as fotos das notícias com SIM são baixadas, reduzidas para até 900 px e **embutidas no HTML**. O arquivo fica com alguns MB, mas abre sem internet, pode ir anexado em e-mail e não fica em branco quando o site bloqueia a exibição da foto fora dele. Quando a matéria não tem foto própria (por exemplo, o site só publica o logo), aparece o painel azul com o ícone do tema.

### Versão resumida (1 página)

O botão laranja **📄 Versão resumida (1 página)**, no canto superior direito, troca a página por uma folha A4 executiva só com o **Top 6** em cards com a foto da matéria e a posição (1 a 6). As duas primeiras ficam em destaque, com foto grande, e as outras quatro com miniatura. Cada card traz o selo de impacto, o tema, o título (clicável), a fonte e a data, além do comentário da equipe quando houver. **Quem entra no Top 6 é você que escolhe**, na coluna **TOP 6** da planilha (antes do STATUS): digite a posição, de 1 a 6. Pode marcar menos de 6, e a página se ajusta. Marcar TOP 6 já leva a notícia para a diretoria, mesmo sem SIM no STATUS. Se a coluna ficar vazia (ou em planilhas antigas, sem essa coluna), o Top 6 é montado sozinho: Alto impacto primeiro e, dentro de cada nível, o maior Score. Se mais de 6 forem marcadas, entram as 6 primeiras posições, e a janela avisa.

O título avisa quantas notícias ficaram só na versão completa ("+ N na versão completa"). Nessa tela, **🖨️ Imprimir / salvar PDF** gera o PDF de uma página para mandar à diretoria, e **← Versão completa** volta. Para abrir o arquivo direto no resumo, acrescente `#resumo` ao endereço.

## Agendamento automático

A tarefa **"Curadoria Omni - Planilha diaria"** do Agendador de Tarefas do Windows roda `1 - Gerar planilha.bat auto` **todo dia às 10h** e abre a planilha nova no Excel. Se o computador estiver desligado às 10h, ela roda assim que ele ligar. Só roda com o usuário logado. No modo `auto` a janela não espera tecla, e o resultado de cada execução fica em `logs/agendamento.log`.

O HTML da diretoria **não** é agendado, porque depende da revisão: depois de marcar SIM/NÃO e salvar, rode o passo 2.

Para mudar o horário ou desligar: abra o **Agendador de Tarefas**, procure "Curadoria Omni - Planilha diaria" e edite ou desabilite a tarefa.

## Arquivos

| Arquivo | Para que serve |
|---|---|
| `motor_pesquisa.py` | Cópia do motor de busca do projeto original (RSS do Google Notícias, notas das fontes, score) |
| `imagens.py` | Converte o link do Google Notícias no endereço real da matéria, lê a foto de destaque do site e a embute reduzida no HTML |
| `gerar_planilha.py` | Etapa 1: pesquisa e monta a planilha (buscas, sugestões de impacto e de área) |
| `gerar_html_diretoria.py` | Etapa 2: lê a planilha e gera o HTML executivo |
| `assets/` | Logos da Omni |
| `planilhas/`, `html_diretoria/` | Saídas geradas |

## Linha de comando

```powershell
.venv\Scripts\python.exe gerar_planilha.py --days 30 --min-credibility 5
.venv\Scripts\python.exe gerar_html_diretoria.py "planilhas\minha_planilha_revisada.xlsx"
```

Instalar dependências num ambiente novo: `python -m venv .venv` e depois `.venv\Scripts\python.exe -m pip install -r requirements.txt`.
