# 💌 Convite Especial — Tati & Eron

Convite web romântico e elegante para o aniversário da **Tati e do Eron**, com:

- 🎯 Personalização individual: cada convidado informa **nome e sobrenome** ao abrir a página, e o convite é gerado especialmente para essa pessoa.
- 💌 Apresentação cinematográfica com tipografia clássica, tema noite-dourado, animações suaves e **chuva de confetes** ao abrir o convite.
- ⏳ Contagem regressiva para o grande dia.
- ✅ **Confirmação de presença (RSVP)** com número de acompanhantes e recadinho carinhoso.
- 📊 Geração automática de **`lista_de_presenca.xlsx`** com os nomes dos convidados confirmados em **ordem alfabética**, atualizada a cada nova confirmação.

## 🚀 Como rodar

Pré-requisito: Node.js 18+.

```bash
npm install
npm start
```

Depois acesse <http://localhost:3000>.

## 🌐 Publicação no GitHub Pages

O convite também pode ser usado **somente como página estática** (sem
backend), publicado automaticamente no GitHub Pages a partir da pasta
[`public/`](./public).

1. No GitHub, abra **Settings → Pages** e em *Build and deployment*
   selecione **Source: GitHub Actions**.
2. Faça push para a branch `main`. O workflow
   [`.github/workflows/pages.yml`](./.github/workflows/pages.yml) publica
   o conteúdo de `public/` no Pages.
3. Acesse `https://<seu-usuario>.github.io/<seu-repo>/` e o convite já
   aparece com tudo: capa personalizada, contagem regressiva, RSVP e
   chuva de confetes.

> 💡 No modo estático (GitHub Pages) não há backend para gravar o
> `lista_de_presenca.xlsx`. As confirmações enviadas pelo formulário
> são registradas no `localStorage` do dispositivo do convidado e
> exibem a mensagem de agradecimento normalmente. Se você rodar o
> servidor Node (`npm start`), o RSVP volta a ser persistido no
> arquivo XLSX automaticamente.

## 📋 Lista de presença (modo servidor)

A cada confirmação de presença, o arquivo **`lista_de_presenca.xlsx`** é criado/atualizado na raiz do projeto, sempre **ordenado alfabeticamente** (sem diferenciar maiúsculas/minúsculas nem acentos), contendo as colunas:

| Nome Completo | Acompanhantes | Mensagem | Confirmado em |
|---------------|---------------|----------|---------------|

## 🗂 Estrutura

```
.
├── server.js                  # Backend (Express) + persistência XLSX
├── package.json
├── lista_de_presenca.xlsx     # Gerado automaticamente
└── public/
    ├── index.html             # Convite + modal de boas-vindas
    ├── styles.css             # Tema visual (glassmorphism, dourado champagne)
    └── script.js              # Lógica: nome do convidado, contagem, RSVP, confetes
```

## ✏️ Personalizando

- **Data e local da festa:** edite os blocos `.detail` em `public/index.html`.
- **Data alvo da contagem regressiva:** ajuste `target` no final de `public/script.js`.
- **Cores e fontes:** todas as variáveis estão no topo de `public/styles.css` (`:root`).

## 🔒 Notas

- O backend valida e sanitiza nome, sobrenome e mensagem (limites de tamanho, controle de duplicidade por nome completo, normalização para evitar XSS).
- O arquivo XLSX é gravado de forma atômica (escreve em `.tmp` e renomeia) e as confirmações são serializadas para evitar corrupção.

Feito com 💜 para um aniversário inesquecível.
