/* ========================================================================
   Convite Especial — Configuração de RSVP via Google Forms
   ------------------------------------------------------------------------
   Para que cada confirmação caia automaticamente em uma planilha do Google
   Sheets, basta criar um Google Form e preencher os campos abaixo. Veja o
   passo a passo detalhado no README.md (seção "RSVP via Google Forms").

   Enquanto os campos estiverem em branco, o convite continua funcionando
   normalmente: usa o backend Express (quando rodando via `npm start`) ou
   grava no `localStorage` do dispositivo (quando hospedado no GitHub
   Pages). Ou seja, é seguro publicar este arquivo "vazio".
   ======================================================================== */

window.CONVITE_CONFIG = {
  googleForm: {
    // URL do endpoint "formResponse" do seu Google Form. Para obtê-la,
    // abra o formulário no modo de visualização (👁), copie a URL da
    // barra de endereço e troque "/viewform..." por "/formResponse".
    // Exemplo:
    //   https://docs.google.com/forms/d/e/1FAIpQLSc.../formResponse
    formResponseUrl: '',

    // IDs dos campos no formato "entry.<numero>". Para descobri-los,
    // use o botão "Obter link pré-preenchido" no menu (⋮) do Google Forms,
    // preencha cada campo com algo distinto, copie o link gerado e leia
    // os pares "entry.123456789=valor" da query string.
    entries: {
      nome: '',                // ex.: 'entry.111111111'
      sobrenome: '',           // ex.: 'entry.222222222'
      acompanhantes: '',       // ex.: 'entry.333333333'
      acompanhantesNomes: '',  // ex.: 'entry.444444444'
      mensagem: '',            // ex.: 'entry.555555555'
    },
  },
};
