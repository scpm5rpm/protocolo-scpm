# Comece aqui (10 minutos)

Este painel é um **suplemento do Excel**: ele abre dentro da planilha e grava direto nela, sem
alterar fórmulas, estilos ou os lançamentos que já existem.

Antes de usar, ele precisa ser **publicado na internet** (uma vez só). O motivo é simples: o Excel
só aceita suplemento que venha de um endereço `https://`. Um arquivo `.html` guardado no OneDrive é
baixado pelo navegador, não executado.

Nada da planilha vai para a internet. Só sobem os arquivos desta pasta (`index.html`, `commands.html`,
`manifest.xml`, `css`, `js`, `assets`).

---

## 1. Criar a conta no GitHub (gratuita)

1. Abra <https://github.com/signup> no Edge.
2. Use o e-mail da seção (de preferência um e-mail que todos possam acessar) e guarde a senha em
   local seguro.
3. Anote o **nome de usuário** escolhido (ex.: `scpm5rpm`). Ele aparece no endereço final.

> O plano gratuito publica apenas repositório **público**. Como aqui só ficam HTML/JS/CSS do painel
> (nenhum dado, nenhuma planilha), isso é aceitável. **Nunca** envie o `.xlsx` para lá.

---

## 2. Apontar o manifesto para o seu endereço

No PowerShell, dentro desta pasta (`ProtocoloApp`):

```powershell
powershell -ExecutionPolicy Bypass -File .\ajustar-endereco.ps1 -Usuario SEUUSUARIO
```

Troque `SEUUSUARIO` pelo nome de usuário do passo 1. O script deve responder:

```
manifest.xml atualizado (10 ocorrencia(s)).
Endereco do painel agora e: https://SEUUSUARIO.github.io/protocolo-scpm
```

Se você rodar de novo, ele apenas confirma que já está certo (não duplica nada).

---

## 3. Criar o repositório e enviar os arquivos

1. Em <https://github.com>, clique em **+** (canto superior direito) → **New repository**.
2. Nome: `protocolo-scpm` (exatamente assim). Visibilidade: **Public**.
   Não marque "Add a README file". Clique em **Create repository**.
3. Na página que abre, clique em **uploading an existing file** (ou **Add file → Upload files**).
4. No Explorer, abra esta pasta (`ProtocoloApp`) **junto com** a página do GitHub. Selecione
   **todo o conteúdo** (Ctrl+A) — menos a planilha, que não está aqui — e **arraste** para a área
   de upload do GitHub. As subpastas `css`, `js` e `assets` entram com a estrutura.
5. Em "Commit changes", clique em **Commit changes** (mensagem pode ficar a padrão).

Confira na lista de arquivos do repositório: `index.html`, `commands.html`, `manifest.xml`, `css/`,
`js/`, `assets/`.

---

## 4. Ligar o site (GitHub Pages)

1. No repositório: **Settings** (engrenagem) → **Pages** (menu da esquerda).
2. Em *Build and deployment*:
   - **Source**: `Deploy from a branch`
   - **Branch**: `main` e pasta `/ (root)`
3. Clique em **Save** e aguarde 1 a 3 minutos (a página recarrega sozinha).
4. No topo do Pages aparece: *Your site is live at* `https://SEUUSUARIO.github.io/protocolo-scpm/`.

Abra esse endereço + `/index.html` no Edge. Deve aparecer o painel com a faixa "Fora do Excel" e o
aviso de que ele só funciona dentro do Excel — **isso é o esperado** fora da planilha.

> Se der **404**: espere mais um pouco e recarregue. Se persistir, confira se o repositório é
> público e se o `index.html` está na **raiz** (não dentro de uma subpasta).

---

## 5. Instalar dentro do Excel Online (cada usuário, uma vez)

1. Abra a planilha `PROTOCOLO DE DOCS E MANIFESTAÇÕES 1.6.2.xlsx` no OneDrive com **Excel Online**
   (botão *Editar no Navegador*).
2. **Página Inicial → Suplementos → Mais Suplementos**.
3. Aba **Meus Suplementos** → **Carregar Meu Suplemento** (*Upload My Add-in*).
4. Escolha o `manifest.xml` desta pasta do OneDrive e confirme.
5. Na aba **Página Inicial** aparece o botão **Protocolo SCPM**. Clique nele: o painel abre à direita.

O carregamento fica salvo na conta Microsoft: em outro computador, com a mesma conta, o botão já
aparece.

---

## 6. Testar numa cópia da planilha (importante)

Não teste na planilha oficial. No OneDrive:

1. Clique com o botão direito no arquivo → **Copiar** → cole na mesma pasta → renomeie para
   `PROTOCOLO - TESTE.xlsx`.
2. Abra a cópia no Excel Online e siga o checklist do `LEIA-ME.md` (item 11).

O que precisa ser conferido no teste real (é o que ainda não foi possível validar aqui):

- a gravação acontece mesmo nas células liberadas da aba protegida;
- a linha nova fica com a mesma aparência das anteriores (bordas, altura);
- as fórmulas das colunas N, R e S calculam sozinhas;
- dois usuários lançando ao mesmo tempo caem em linhas diferentes.

---

## Se algo der errado

| O que aparece | O que fazer |
|---|---|
| Página 404 no endereço do Pages | aguarde 3 min; confira repositório público e `index.html` na raiz |
| "Abra este painel dentro do Excel Online" | você abriu o endereço no navegador, fora do Excel — abra a planilha e clique em **Protocolo SCPM** |
| Botão do suplemento não aparece | refaça o passo 5; confira se o `manifest.xml` enviado é o **já ajustado** (sem `SEUUSUARIO`) |
| "A planilha está protegida e a célula de destino não permite edição" | o intervalo liberado da aba não cobre a linha — veja `LEIA-ME.md` item 6 |
| "A aba esperada não foi encontrada" | abra o arquivo `PROTOCOLO DE DOCS E MANIFESTAÇÕES` (nome exato das abas) |

Detalhes completos, regras da planilha, proteção, backup e diagnóstico: **`LEIA-ME.md`**.
