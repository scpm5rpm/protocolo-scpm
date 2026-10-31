# Protocolo SCPM — painel do Excel

Aplicativo web (suplemento do Excel) para lançar, consultar e **alterar** registros na planilha
`PROTOCOLO DE DOCS E MANIFESTAÇÕES 1.6.2.xlsx`, usando o próprio arquivo como banco de dados.

- Roda **dentro do Excel Online** (Edge/Chrome), aberto pelo OneDrive.
- Vários militares podem usar ao mesmo tempo: o Excel Online combina as edições de todos.
- O painel grava **somente nas colunas de digitação**. Fórmulas, estilos, validações, cabeçalhos e
  os dados já lançados não são alterados.

> **Situação atual:** as abas **Entrada** e **Saída** estão completas (gravar, consultar e alterar)
> e o núcleo já foi testado com um Excel simulado (item 12). Falta publicar a página em um endereço
> HTTPS e executar o checklist do item 11 numa cópia da planilha. Manifestação, Proposta de
> Recompensa e Controle CEDMU estão mapeadas e virão depois (item 10).

---

## 1. Estrutura da pasta

```
ProtocoloApp/
├─ COMECE-AQUI.md         ← passo a passo de publicação e instalação (comece por ele)
├─ manifest.xml          ← manifesto do suplemento (é o arquivo que se carrega no Excel)
├─ index.html            ← painel (formulários, consulta, ajuda)
├─ commands.html         ← página exigida pelo manifesto (não tem interface)
├─ ajustar-endereco.ps1  ← troca o endereço de exemplo do manifesto pelo endereço real
├─ css/styles.css
├─ js/schema.js          ← mapa de abas, colunas, fórmulas e listas
├─ js/excel.js           ← leitura/gravação no arquivo via Office.js
├─ js/app.js             ← telas e regras
└─ assets/icon-16.png | icon-32.png | icon-80.png
```

Guarde a pasta **junto com a planilha**, na mesma pasta do OneDrive. O `manifest.xml` precisa
estar acessível a cada usuário na hora de instalar o suplemento.

---

## 2. Publicar o painel (uma vez só)

### 2.1. Por que o painel não roda direto do OneDrive

A pasta `ProtocoloApp` fica **junto da planilha, no OneDrive** — é dela que cada usuário carrega o
`manifest.xml`. A página do painel, porém, **não pode ser executada de dentro do OneDrive**:

- quando o Edge abre um `.html` armazenado no OneDrive, ele **baixa** o arquivo em vez de executar;
- o Excel exige que a página do suplemento venha de um endereço **HTTPS público** (o `localhost`
  serve só para desenvolvimento em uma máquina).

Nada da planilha vai para a hospedagem: o repositório/pasta pública guarda apenas HTML, JS, CSS e
ícones do painel. A planilha continua exclusivamente no OneDrive.

Opções, da mais simples para a mais estável:

| Opção | Precisa de | Observação |
|---|---|---|
| **GitHub Pages** (recomendado) | conta gratuita no GitHub | passo a passo abaixo |
| Pasta pública HTTPS da PMMG/5ª RPM | apoio da área de TI | melhor opção se existir; use `ajustar-endereco.ps1 -Endereco` |
| SharePoint/OneDrive institucional | conta institucional Microsoft 365 | hoje não disponível para a seção |
| Outro servidor HTTPS estático | acesso ao servidor | qualquer um serve, desde que aceite os arquivos da pasta |

Se a seção conseguir uma pasta institucional, basta publicar o conteúdo de `ProtocoloApp` nela e
rodar:

```powershell
powershell -ExecutionPolicy Bypass -File .\ajustar-endereco.ps1 -Endereco https://endereco.da.secao/protocolo
```

### 2.2. Criar o repositório e enviar os arquivos

> Atalho: o `COMECE-AQUI.md` tem o mesmo passo a passo em formato de checklist.

1. Crie uma conta em <https://github.com> (se ainda não tiver) e anote o nome de usuário.
2. **Antes de enviar**, rode o script do item 2.3 para o manifesto já subir com o endereço certo.
3. Crie um repositório **público** chamado exatamente `protocolo-scpm` (*New repository*, sem README).
   - Precisa ser público: o GitHub Pages gratuito não publica repositório privado.
   - O repositório contém apenas o painel (HTML/JS/CSS), **nunca a planilha**.
4. Envie o conteúdo da pasta `ProtocoloApp` para a raiz do repositório:
   - pelo site: **Add file → Upload files** e arraste **o conteúdo** da pasta (os arquivos e as
     subpastas `css`, `js`, `assets`); ou
   - pelo Git: `git init`, `git add .`, `git commit -m "painel"`, `git push`.

### 2.3. Apontar o manifesto para o endereço real

No arquivo `manifest.xml`, todas as ocorrências de `https://SEUUSUARIO.github.io/protocolo-scpm`
precisam virar o seu endereço. Faça de uma das duas formas:

**Opção A — script (mais seguro).** No PowerShell, dentro da pasta do suplemento:

```powershell
# GitHub Pages
powershell -ExecutionPolicy Bypass -File .\ajustar-endereco.ps1 -Usuario SEUUSUARIO

# Hospedagem institucional (ou qualquer outra URL base)
powershell -ExecutionPolicy Bypass -File .\ajustar-endereco.ps1 -Endereco https://endereco.da.secao/protocolo
```

O script avisa quantas ocorrências trocou, pode ser rodado de novo sem duplicar nada e mantém o
arquivo em UTF-8 sem BOM.

**Opção B — manual.** Abra `manifest.xml` e troque as 10 ocorrências de `SEUUSUARIO`.

Depois de alterar, **envie o `manifest.xml` atualizado para o repositório** (se ainda não enviou, ele
já sobe ajustado) e deixe essa versão na pasta do OneDrive — é ela que os usuários carregam no Excel.

### 2.4. Ligar o GitHub Pages e testar

1. No repositório: **Settings → Pages**.
2. Em *Build and deployment*: **Source = Deploy from a branch**, **Branch = main**, **pasta = / (root)**.
3. Salve e aguarde de 1 a 3 minutos.
4. Teste no navegador: `https://SEUUSUARIO.github.io/protocolo-scpm/index.html`
   Deve aparecer o painel com a mensagem "só funciona dentro do Excel" (é o esperado fora do Excel).
   Se der 404, espere mais um pouco; confira se o repositório é público e se `index.html` está na raiz.

---

## 3. Instalar o suplemento (cada usuário, uma vez)

### Excel Online (recomendado)

1. Abra a planilha no OneDrive com **Excel Online** (botão *Editar no Navegador*).
2. **Página Inicial → Suplementos → Mais Suplementos**.
3. Na janela *Suplementos do Office*, aba **Meus Suplementos**, clique em
   **Carregar Meu Suplemento** (*Upload My Add-in*).
4. Escolha o `manifest.xml` que está na pasta do OneDrive e confirme.
5. O botão **Protocolo SCPM** aparece na faixa de opções, aba **Página Inicial**. Clique nele.

> O carregamento fica salvo na sua conta Microsoft. Em outro computador, com a mesma conta, o
> suplemento já aparece — não precisa repetir.

### Excel para desktop (opcional)

1. **Arquivo → Opções → Central de Confiabilidade → Configurações da Central de Confiabilidade →
   Catálogos de Suplementos Confiáveis**.
2. Adicione a pasta do OneDrive que contém o `manifest.xml` e marque *Mostrar no Menu*.
3. **Inserir → Meus Suplementos → Pasta Compartilhada → Protocolo SCPM**.

O suplemento foi pensado para o Excel Online. No desktop ele funciona, mas a gravação simultânea
é menos fluida.

---

## 4. Uso diário

### Lançar uma entrada

1. Aba **Entrada** no painel.
2. Preencha a data, o tipo de documento (lista) e os demais campos conhecidos.
3. **Gravar entrada** (ou `Ctrl + Enter`).
4. O painel mostra a **linha** e o **protocolo** gerados. Use **Ver na planilha** para conferir.

### Lançar uma saída

Mesma coisa na aba **Saída**. A data do documento é obrigatória porque é ela que localiza a
próxima linha livre.

### Consultar

Aba **Consulta**: pesquisa por protocolo, nº do documento, tipo, envolvido, origem, destino etc.
O botão **Atualizar** relê a planilha (os dados ficam em cache enquanto você navega).

### Alterar um lançamento

1. Aba **Consulta** → localize o registro.
2. Clique em **Editar**. O formulário abre preenchido com os dados daquela linha.
3. Altere o que precisar e clique em **Salvar alterações**.
4. Para desistir: **Cancelar**, **Limpar** ou tecla `Esc`.

Se outra pessoa alterou a mesma linha enquanto você editava, o painel avisa **antes** de
sobrescrever e pergunta se você quer gravar mesmo assim.

---

## 5. O que o painel altera (e o que não altera)

| Aba | Colunas em que o painel escreve | Colunas com fórmula (intocadas) |
|---|---|---|
| Entrada de Documentos | C, H, I, J, K, L, M, O, P, Q, T, U, V, W, X, Y | A, B, D, E, F, G, N, R, S |
| Saída e Produção | B, D, I, J, K, L, M, N, O, P, Q, T, U, V, W, X, Y | A, C, E, F, G, H, R, S |

Regras fixas:

- Gravação sempre na **próxima linha livre** (a planilha tem os protocolos pré-preenchidos).
- O **protocolo (coluna A)** é lido da planilha, nunca escrito pelo painel.
- Não há inserção/remoção de linhas, colunas ou abas — a estrutura fica intacta.
- Ao alterar, o painel reescreve apenas as colunas acima, na linha já existente.
- **Datas** são gravadas como data do Excel (número de série), sem risco de o dia "pular" por fuso
  horário. Se a célula de destino estiver no formato *Geral*, o painel aplica `dd/mm/aaaa` **apenas
  naquela célula**; se já houver formato de data, ele é mantido.
- A **data do documento é obrigatória**: é ela que localiza a próxima linha livre. Datas
  impossíveis (ex.: 30/02) são recusadas com aviso, para não criar linha "invisível" que seria
  sobrescrita no lançamento seguinte.
- Na alteração, um campo deixado em branco no formulário **limpa** aquela célula na planilha.
  Campos que o formulário não envia não são tocados.

---

## 6. Proteção das abas e limite de linhas

As abas **Entrada** e **Saída** são protegidas com senha, mas têm *intervalos liberados para
edição* (Permitir aos usuários editar intervalos). O painel escreve justamente nesses intervalos.

Hoje os intervalos liberados alcançam:

| Aba | Até a linha | Observação |
|---|---|---|
| Entrada de Documentos | **15.170** | a partir de 15.171 a aba fica totalmente bloqueada |
| Saída e Produção | **16.194** | a partir de 16.195, idem |

Próximas linhas livres no momento em que este documento foi escrito: **Entrada 14.584** e
**Saída 15.310**.

### Se aparecer "A planilha está protegida e a célula de destino não permite edição"

1. **Confirme se o problema é só na linha de destino.** O painel já avisa quando a próxima linha
   passa do limite liberado.
2. **Com a senha em mãos** (quem mantém a planilha):
   - Aba **Revisão → Desproteger Planilha** na aba correspondente.
   - Selecione da linha 1 até a última linha prevista e libere as colunas de digitação:
     - Entrada: **C** e **H:Y**
     - Saída: **B**, **D** e **I:Y**
   - **Revisão → Proteger Planilha** novamente (pode manter a mesma senha).
3. **Sem a senha:** não é possível alterar a proteção. Nesse caso, o caminho é criar uma cópia
   nova da planilha (o que muda o histórico de fórmulas) ou obter a senha com quem configurou.
   Não tente "quebrar" a senha no arquivo em produção.

> Importante: nunca trabalhe direto na planilha oficial sem backup (item 7).

---

## 7. Backup e recuperação

- **Histórico de versões do OneDrive (principal):** no site do OneDrive, clique com o botão
  direito no arquivo → **Histórico de Versões**. Qualquer gravação errada pode ser desfeita.
- **Cópia diária/semanal:** antes de períodos de muito lançamento, faça
  *Arquivo → Salvar como → Cópia* com a data no nome
  (`PROTOCOLO ... - BACKUP 2026-10-30.xlsx`).
- Combine com a seção quem pode restaurar versões, para não depender de uma só pessoa.

---

## 8. Problemas comuns

| Sintoma | Causa provável | O que fazer |
|---|---|---|
| "Abra este painel dentro do Excel Online" | página aberta no navegador, fora do Excel | abra a planilha no Excel Online e clique em **Protocolo SCPM** |
| Botão do suplemento não aparece | manifesto não carregado ou endereço errado | refaça o item 3; confira se trocou `SEUUSUARIO` |
| "A aba esperada não foi encontrada" | planilha errada aberta | abra o arquivo `PROTOCOLO DE DOCS E MANIFESTAÇÕES` |
| "Não foi possível reservar uma linha livre" | duas pessoas gravaram no mesmo instante | tente de novo; a planilha não foi corrompida |
| Lista de tipos vazia | aba `LISTAS` alterada | avise o responsável técnico; os campos texto continuam aceitando digitação |
| Linha errada ao alterar | cache da consulta antigo | clique em **Atualizar** e localize de novo |

Diagnóstico rápido: aba **Ajuda → Diagnóstico** mostra o nome do arquivo, a próxima linha livre de
cada aba, se a aba está protegida e o limite liberado.

---

## 9. Como o painel encontra a linha certa (referência técnica)

| Aba | Regra |
|---|---|
| Entrada de Documentos | próxima linha sem data na coluna C |
| Saída e Produção | próxima linha sem data na coluna D |
| Dados da Manifestação | linha da Entrada **− 1** |
| Proposta de Recompensa | mesma linha da Entrada |
| Controle CEDMU | linha da Saída **− 1.474** |

Cuidados já embutidos no código:

- Nome real da aba de saída é `" Saída e Produção"`, **com um espaço no início**.
- Células de data voltam do Office.js como número de série (época 1899-12-30); a conversão já é
  feita no painel, tanto na leitura quanto na conferência.
- A gravação relê a linha depois de escrever e confere se o conteúdo é o esperado; se outra pessoa
  ocupou a linha, o painel tenta de novo (até 3 vezes) e informa a linha reservada.
- Se a disputa acontecer **durante** a escrita, o painel grava na linha seguinte e mostra um aviso
  pedindo para conferir a linha que ficou misturada.
- A alteração compara os valores lidos no início da edição com o estado atual da linha antes de
  gravar; se alguém mudou algo no meio do caminho, o painel pergunta antes de sobrescrever.
- Se o campo data do documento ficar vazio (ou inválido) numa alteração, o painel recusa: sem a
  data a linha deixaria de ser encontrada e poderia ser sobrescrita depois.

---

## 10. Fase 2 — formulários complementares

Já mapeados, aguardando validação do núcleo Entrada/Saída:

- **Dados da Manifestação** (colunas AF:DW, linha da Entrada − 1): dados do reclamante, do fato,
  classificação, natureza e encaminhamentos.
- **Proposta de Recompensa** (colunas J:BX, exceto fórmulas, mesma linha da Entrada).
- **Controle CEDMU** (colunas L:AG, exceto fórmulas, linha da Saída − 1.474).

Antes de liberar essas abas é preciso testar em uma **cópia** da planilha, porque essas abas não
têm proteção e qualquer erro de mapeamento grava na coluna errada.

---

## 11. Checklist de teste (fazer em uma cópia, nunca na oficial)

1. Copie o arquivo e abra a cópia no Excel Online.
2. Instale o suplemento e abra o painel.
3. Em **Ajuda → Diagnóstico**, confira arquivo, próxima linha e proteção.
4. Lance **1 entrada** e confira: linha correta, protocolo preenchido, fórmulas das colunas
   N, R e S calculadas, estilo da linha preservado.
5. Lance **1 saída** e confira o mesmo.
6. Na **Consulta**, edite a entrada lançada, mude um campo e salve; confira na planilha.
7. Abra a planilha em duas janelas (ou dois usuários) e lance ao mesmo tempo; confira se as duas
   linhas ficaram corretas.
8. Feche e reabra o Excel Online; confira se o suplemento continua instalado.

Só depois desse checklist aponte o painel para a planilha oficial.

---

## 12. Situação dos testes (30/10/2026)

Já validado com um **Excel simulado** que reproduz a API do Office.js (mesmas chamadas, datas em
número de série, abas protegidas, listas da aba `LISTAS`):

- formulários da Entrada (16 campos) e da Saída (17 campos) gerados a partir do mapa de colunas;
- 13 listas de sugestão carregadas da aba `LISTAS`;
- gravação na próxima linha livre, com protocolo lido da planilha;
- data gravada como número de série e formato `dd/mm/aaaa` aplicado só quando a célula está *Geral*;
- recusa de data vazia/inválida e de linha além do trecho liberado (15.170 / 16.194);
- leitura das últimas linhas e montagem da Consulta;
- alteração de linha existente, limpeza de campo e detecção de conflito com confirmação;
- dois lançamentos simultâneos: o segundo foi para a linha seguinte (tentativa 2), sem misturar;
- diagnóstico (arquivo, próxima linha, proteção, limite) e mensagem ao abrir fora do Excel.

Ainda **não** validado em Excel de verdade — é o próximo passo, sempre numa **cópia** da planilha:

- gravação em célula liberada de aba protegida (o ponto de maior risco do projeto);
- aparência da linha nova (bordas, altura, formatação) e recálculo das fórmulas N, R e S;
- instalação por *Carregar Meu Suplemento* na conta Microsoft usada pela seção;
- comportamento com dois usuários reais, ao mesmo tempo, no Excel Online.
