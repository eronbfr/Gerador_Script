/* ========================================================================
   Convite Especial — Tati & Eron
   Lógica do front-end: boas-vindas, personalização, contagem regressiva,
   RSVP e confetes.
   ======================================================================== */

(function () {
  'use strict';

  // -------- Utilidades --------

  function $(sel, root) { return (root || document).querySelector(sel); }
  function $$(sel, root) { return Array.from((root || document).querySelectorAll(sel)); }

  function escapeHTML(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function toTitleCasePt(str) {
    return String(str || '')
      .trim()
      .replace(/\s+/g, ' ')
      .toLowerCase()
      .split(' ')
      .map(function (p) { return p ? p.charAt(0).toLocaleUpperCase('pt-BR') + p.slice(1) : p; })
      .join(' ');
  }

  // -------- Confetes (canvas leve, sem dependências) --------

  function startConfetti(durationMs) {
    var canvas = document.getElementById('confetti-canvas');
    if (!canvas) return;
    var prefersReduced = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    if (prefersReduced) return;

    var ctx = canvas.getContext('2d');
    var dpr = window.devicePixelRatio || 1;
    function resize() {
      canvas.width = window.innerWidth * dpr;
      canvas.height = window.innerHeight * dpr;
      canvas.style.width = window.innerWidth + 'px';
      canvas.style.height = window.innerHeight + 'px';
      ctx.setTransform(1, 0, 0, 1, 0, 0);
      ctx.scale(dpr, dpr);
    }
    resize();
    window.addEventListener('resize', resize);

    var colors = ['#e9c98b', '#f8e3b0', '#f4a3c6', '#c95a90', '#ffffff', '#b88a3d'];
    var pieces = [];
    var count = Math.min(180, Math.floor(window.innerWidth / 6));
    for (var i = 0; i < count; i++) {
      pieces.push({
        x: Math.random() * window.innerWidth,
        y: Math.random() * -window.innerHeight,
        w: 6 + Math.random() * 6,
        h: 8 + Math.random() * 8,
        color: colors[Math.floor(Math.random() * colors.length)],
        rot: Math.random() * Math.PI * 2,
        vRot: (Math.random() - 0.5) * 0.2,
        vY: 1.5 + Math.random() * 2.5,
        vX: (Math.random() - 0.5) * 1.5,
        shape: Math.random() < 0.5 ? 'rect' : 'circle',
      });
    }

    var start = performance.now();
    var rafId;
    function frame(now) {
      var elapsed = now - start;
      ctx.clearRect(0, 0, window.innerWidth, window.innerHeight);
      pieces.forEach(function (p) {
        p.x += p.vX;
        p.y += p.vY;
        p.rot += p.vRot;
        if (p.y > window.innerHeight + 20) {
          p.y = -20;
          p.x = Math.random() * window.innerWidth;
        }
        ctx.save();
        ctx.translate(p.x, p.y);
        ctx.rotate(p.rot);
        ctx.fillStyle = p.color;
        ctx.globalAlpha = Math.max(0, 1 - elapsed / durationMs);
        if (p.shape === 'rect') {
          ctx.fillRect(-p.w / 2, -p.h / 2, p.w, p.h);
        } else {
          ctx.beginPath();
          ctx.arc(0, 0, p.w / 2, 0, Math.PI * 2);
          ctx.fill();
        }
        ctx.restore();
      });

      if (elapsed < durationMs) {
        rafId = requestAnimationFrame(frame);
      } else {
        ctx.clearRect(0, 0, window.innerWidth, window.innerHeight);
        cancelAnimationFrame(rafId);
      }
    }
    rafId = requestAnimationFrame(frame);
  }

  // -------- Contagem regressiva --------

  function startCountdown(targetDate) {
    var els = {
      days: $('[data-unit="days"]'),
      hours: $('[data-unit="hours"]'),
      minutes: $('[data-unit="minutes"]'),
      seconds: $('[data-unit="seconds"]'),
    };
    function pad(n) { return String(Math.max(0, n)).padStart(2, '0'); }
    function tick() {
      var diff = targetDate.getTime() - Date.now();
      if (diff <= 0) {
        els.days.textContent = '00';
        els.hours.textContent = '00';
        els.minutes.textContent = '00';
        els.seconds.textContent = '00';
        return;
      }
      var s = Math.floor(diff / 1000);
      var d = Math.floor(s / 86400);
      var h = Math.floor((s % 86400) / 3600);
      var m = Math.floor((s % 3600) / 60);
      var sec = s % 60;
      els.days.textContent = pad(d);
      els.hours.textContent = pad(h);
      els.minutes.textContent = pad(m);
      els.seconds.textContent = pad(sec);
    }
    tick();
    setInterval(tick, 1000);
  }

  // -------- Personalização do convite --------

  function personalizeInvite(nome, sobrenome) {
    var nomeCap = toTitleCasePt(nome);
    var sobrenomeCap = toTitleCasePt(sobrenome);
    var nomeCompleto = (nomeCap + ' ' + sobrenomeCap).trim();
    $('#guest-name').textContent = nomeCompleto;
    $('#guest-greeting').textContent = 'Querido(a) ' + nomeCap + ',';
    document.title = '💌 Convite para ' + nomeCompleto + ' — Tati & Eron';
    try {
      sessionStorage.setItem('convite.nome', nomeCap);
      sessionStorage.setItem('convite.sobrenome', sobrenomeCap);
    } catch (_) { /* ignora */ }
  }

  // -------- Modal de boas-vindas --------

  function setupWelcome() {
    var overlay = $('#welcome-overlay');
    var form = $('#welcome-form');
    var inputNome = $('#input-nome');
    var inputSobrenome = $('#input-sobrenome');
    var errorEl = $('#welcome-error');
    var invite = $('#invite');

    // Pré-preenche se já houver dados na sessão.
    try {
      var savedNome = sessionStorage.getItem('convite.nome');
      var savedSobrenome = sessionStorage.getItem('convite.sobrenome');
      if (savedNome) inputNome.value = savedNome;
      if (savedSobrenome) inputSobrenome.value = savedSobrenome;
    } catch (_) { /* ignora */ }

    setTimeout(function () { inputNome.focus(); }, 400);

    form.addEventListener('submit', function (ev) {
      ev.preventDefault();
      errorEl.textContent = '';

      var nome = (inputNome.value || '').trim();
      var sobrenome = (inputSobrenome.value || '').trim();
      var nameRegex = /^[A-Za-zÀ-ÖØ-öø-ÿ' -]{2,40}$/;

      if (!nameRegex.test(nome)) {
        errorEl.textContent = 'Por favor, informe um nome válido (apenas letras).';
        inputNome.focus();
        return;
      }
      if (!nameRegex.test(sobrenome)) {
        errorEl.textContent = 'Por favor, informe um sobrenome válido (apenas letras).';
        inputSobrenome.focus();
        return;
      }

      personalizeInvite(nome, sobrenome);

      overlay.classList.add('closing');
      setTimeout(function () {
        overlay.style.display = 'none';
        invite.classList.remove('hidden');
        invite.setAttribute('aria-hidden', 'false');
        startConfetti(4500);
        // Foca no convite para leitores de tela.
        invite.setAttribute('tabindex', '-1');
        invite.focus({ preventScroll: true });
      }, 480);
    });
  }

  // -------- Formulário de RSVP --------

  function setupRSVP() {
    var form = $('#rsvp-form');
    var extra = $('#extra-fields');
    var errorEl = $('#rsvp-error');
    var submitBtn = $('#rsvp-submit');
    var success = $('#rsvp-success');
    var successTitle = $('#rsvp-success-title');
    var successMsg = $('#rsvp-success-msg');
    var successCounter = $('#rsvp-success-counter');

    // Mostra/esconde campos extras conforme escolha.
    $$('input[name="presenca"]').forEach(function (radio) {
      radio.addEventListener('change', function () {
        if (radio.value === 'sim' && radio.checked) {
          extra.hidden = false;
        } else if (radio.value === 'nao' && radio.checked) {
          extra.hidden = true;
        }
      });
    });

    form.addEventListener('submit', async function (ev) {
      ev.preventDefault();
      errorEl.textContent = '';

      var presenca = (form.querySelector('input[name="presenca"]:checked') || {}).value;
      if (!presenca) {
        errorEl.textContent = 'Por favor, escolha uma das opções.';
        return;
      }

      var nome = '';
      var sobrenome = '';
      try {
        nome = sessionStorage.getItem('convite.nome') || '';
        sobrenome = sessionStorage.getItem('convite.sobrenome') || '';
      } catch (_) { /* ignora */ }

      if (!nome || !sobrenome) {
        errorEl.textContent = 'Recarregue a página e informe seu nome novamente.';
        return;
      }

      // Resposta "não vou" — agradecemos sem chamar o servidor.
      if (presenca === 'nao') {
        form.hidden = true;
        success.hidden = false;
        successTitle.textContent = 'Sentiremos sua falta 💔';
        successMsg.innerHTML = 'Obrigado por avisar, <strong>' +
          escapeHTML(toTitleCasePt(nome)) + '</strong>. Mandamos um abraço apertado!';
        successCounter.textContent = '';
        return;
      }

      var acompanhantes = parseInt($('#acompanhantes').value, 10);
      if (!Number.isFinite(acompanhantes) || acompanhantes < 0) acompanhantes = 0;
      if (acompanhantes > 10) acompanhantes = 10;
      var mensagem = ($('#mensagem').value || '').trim().slice(0, 280);

      submitBtn.disabled = true;
      var originalText = submitBtn.innerHTML;
      submitBtn.innerHTML = '<span>Enviando...</span>';

      var nomeCompleto = (toTitleCasePt(nome) + ' ' + toTitleCasePt(sobrenome)).trim();

      function showSuccess(opts) {
        form.hidden = true;
        success.hidden = false;
        successTitle.textContent = opts.duplicate ? 'Já estava confirmado! 💖' : 'Presença confirmada! 🎉';
        successMsg.innerHTML = escapeHTML(opts.message);
        successCounter.textContent = opts.counter || '';
        startConfetti(3500);
      }

      // Persistência local (usada também como fallback quando não há backend,
      // por exemplo quando o convite é hospedado estaticamente no GitHub Pages).
      function rememberLocally() {
        try {
          var raw = localStorage.getItem('convite.rsvps');
          var list = raw ? JSON.parse(raw) : [];
          if (!Array.isArray(list)) list = [];
          var key = nomeCompleto.toLowerCase();
          var already = list.some(function (g) { return (g.nome || '').toLowerCase() === key; });
          if (!already) {
            list.push({
              nome: nomeCompleto,
              acompanhantes: acompanhantes,
              mensagem: mensagem,
              data: new Date().toISOString(),
            });
            localStorage.setItem('convite.rsvps', JSON.stringify(list));
          }
          return { duplicate: already, total: list.length };
        } catch (_) {
          return { duplicate: false, total: null };
        }
      }

      try {
        var res = await fetch('api/rsvp', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json', 'Accept': 'application/json' },
          body: JSON.stringify({
            nome: nome,
            sobrenome: sobrenome,
            acompanhantes: acompanhantes,
            mensagem: mensagem,
          }),
        });

        // Se o backend não está disponível (ex.: GitHub Pages), o servidor
        // responde com HTML (404). Nesse caso, caímos para o modo offline.
        var contentType = res.headers.get('content-type') || '';
        if (!res.ok || contentType.indexOf('application/json') === -1) {
          throw new Error('offline');
        }
        var data = await res.json();
        if (!data.ok) {
          throw new Error(data.error || 'Não foi possível confirmar agora.');
        }

        rememberLocally();
        showSuccess({
          duplicate: !!data.duplicate,
          message: data.message || 'Obrigado!',
          counter: typeof data.total === 'number' ? 'Você é a confirmação nº ' + data.total + '.' : '',
        });
      } catch (err) {
        // Fallback offline / estático: registra localmente e agradece.
        var local = rememberLocally();
        showSuccess({
          duplicate: local.duplicate,
          message: local.duplicate
            ? 'Já registramos a sua confirmação, ' + nomeCompleto + '! 💖'
            : 'Presença confirmada com sucesso, ' + nomeCompleto + '! 🎉',
          counter: local.total ? 'Você é a confirmação nº ' + local.total + ' (registrada neste dispositivo).' : '',
        });
      }
    });
  }

  // -------- Inicialização --------

  document.addEventListener('DOMContentLoaded', function () {
    setupWelcome();
    setupRSVP();
    // Data alvo padrão: 60 dias a partir de hoje (placeholder até definir a data oficial).
    var target = new Date();
    target.setDate(target.getDate() + 60);
    target.setHours(20, 0, 0, 0);
    startCountdown(target);
  });
})();
