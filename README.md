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
>
> Para centralizar todas as confirmações em uma única planilha sem
> precisar rodar servidor, configure o **RSVP via Google Forms** logo
> abaixo.

## 📨 RSVP via Google Forms (recomendado para o GitHub Pages)

Esta é a forma mais simples de receber as confirmações em **uma única
planilha do Google Sheets**, sem precisar manter servidor algum. Só
precisa fazer uma vez:

1. **Crie o formulário** em <https://forms.google.com> com 5 perguntas
   (todas do tipo *Resposta curta*, exceto a mensagem que pode ser
   *Parágrafo*):

   | # | Pergunta no Form        | Mapeia para               |
   |---|-------------------------|---------------------------|
   | 1 | Nome                    | `entries.nome`            |
   | 2 | Sobrenome               | `entries.sobrenome`       |
   | 3 | Acompanhantes (número)  | `entries.acompanhantes`   |
   | 4 | Nomes dos acompanhantes | `entries.acompanhantesNomes` |
   | 5 | Mensagem                | `entries.mensagem`        |

2. No menu **⋮** do Google Forms, escolha **"Obter link
   pré-preenchido"**, preencha cada pergunta com algo único (ex.:
   `nome=AAA`, `sobrenome=BBB`, `acompanhantes=1`, etc.) e clique em
   **"Obter link"**. Copie a URL gerada — ela contém pares no formato
   `entry.123456789=AAA`. Anote o número de cada `entry` correspondente
   a cada pergunta.

3. **Pegue a URL de envio**: abra o formulário no modo de visualização
   (botão 👁), copie a URL e troque o final `viewform...` por
   `formResponse`. Deve ficar algo como:
   `https://docs.google.com/forms/d/e/1FAIpQLSc.../formResponse`.

4. **Vincule a planilha**: na aba *Respostas* do Form, clique no ícone
   verde do Sheets para criar/escolher a planilha onde as respostas
   serão acumuladas. Pronto — a partir daí cada confirmação aparece
   automaticamente lá, e você pode exportar como `.xlsx` quando
   quiser (*Arquivo → Fazer download → Microsoft Excel*).

5. **Configure o site**: edite [`public/config.js`](./public/config.js)
   com a URL e os IDs `entry.<numero>` que você anotou:

   ```js
   window.CONVITE_CONFIG = {
     googleForm: {
       formResponseUrl: 'https://docs.google.com/forms/d/e/1FAIpQLSc.../formResponse',
       entries: {
         nome: 'entry.111111111',
         sobrenome: 'entry.222222222',
         acompanhantes: 'entry.333333333',
         acompanhantesNomes: 'entry.444444444',
         mensagem: 'entry.555555555',
       },
     },
   };
   ```

   Depois faça push para `main`. Em alguns segundos o GitHub Pages
   republica e o convite passa a enviar cada RSVP direto para o seu
   Google Sheets.

> 🛡️ Enquanto `formResponseUrl` ou `entries.nome` estiverem em branco,
> o convite continua funcionando no modo anterior (servidor + fallback
> em `localStorage`). Por isso é seguro publicar com o `config.js`
> "vazio" enquanto você prepara o formulário.

## 📋 Lista de presença (modo servidor)

A cada confirmação de presença, o arquivo **`lista_de_presenca.xlsx`** é criado/atualizado na raiz do projeto, sempre **ordenado alfabeticamente** (sem diferenciar maiúsculas/minúsculas nem acentos), contendo as colunas:

| Nome Completo | Acompanhantes | Nomes dos Acompanhantes | Mensagem | Confirmado em |
|---------------|---------------|-------------------------|----------|---------------|

## 🗂 Estrutura

```
.
├── server.js                  # Backend (Express) + persistência XLSX
├── package.json
├── lista_de_presenca.xlsx     # Gerado automaticamente
└── public/
    ├── index.html             # Convite + modal de boas-vindas
    ├── styles.css             # Tema visual (glassmorphism, dourado champagne)
    ├── config.js              # Configuração opcional do Google Form (RSVP)
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
