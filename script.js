// ============== SLIDE PRESENTATION with TTS ==============
(function() {
  const container = document.getElementById('slidesContainer');
  if (!container) return;

  const slides = Array.from(document.querySelectorAll('.slide'));
  const slideNumEl = document.getElementById('slideNum');
  const timerEl = document.getElementById('timer');
  const progressFill = document.getElementById('progressFill');
  const prevBtn = document.getElementById('prevBtn');
  const nextBtn = document.getElementById('nextBtn');
  const speakBtn = document.getElementById('speakBtn');
  const stopBtn = document.getElementById('stopBtn');

  const SLIDE_DURATION = 30; // seconds
  let current = 0;
  let remaining = SLIDE_DURATION;
  let timerInterval = null;
  let utterance = null;

  function showSlide(index) {
    if (index < 0 || index >= slides.length) return;
    slides.forEach(s => s.classList.remove('active'));
    slides[index].classList.add('active');
    current = index;
    slideNumEl.textContent = index + 1;
    remaining = SLIDE_DURATION;
    updateTimer();
    stopSpeech();
  }

  function updateTimer() {
    timerEl.textContent = remaining;
    const progress = ((SLIDE_DURATION - remaining) / SLIDE_DURATION) * 100;
    progressFill.style.width = progress + '%';
  }

  function tick() {
    remaining--;
    updateTimer();
    if (remaining <= 0) {
      if (current < slides.length - 1) {
        showSlide(current + 1);
      } else {
        stopTimer();
        progressFill.style.width = '100%';
        timerEl.textContent = 'Done!';
      }
    }
  }

  function startTimer() {
    stopTimer();
    timerInterval = setInterval(tick, 1000);
  }

  function stopTimer() {
    if (timerInterval) clearInterval(timerInterval);
    timerInterval = null;
  }

  // ---- Text-to-Speech ----
  function speakCurrentSlide() {
    if (!('speechSynthesis' in window)) {
      alert('Your browser does not support speech synthesis.');
      return;
    }
    stopSpeech();
    const text = slides[current].getAttribute('data-text') || '';
    if (!text.trim()) return;
    utterance = new SpeechSynthesisUtterance(text);
    utterance.lang = 'en-US';
    utterance.rate = 0.95;
    utterance.pitch = 1.0;

    const voices = speechSynthesis.getVoices();
    const preferred = voices.find(v => v.lang === 'en-US' && /female|samantha|zira|google us/i.test(v.name))
        || voices.find(v => v.lang.startsWith('en-US'))
        || voices.find(v => v.lang.startsWith('en'));
    if (preferred) utterance.voice = preferred;

    speechSynthesis.speak(utterance);
  }

  function stopSpeech() {
    if ('speechSynthesis' in window) {
      speechSynthesis.cancel();
    }
  }

  prevBtn.addEventListener('click', () => {
    if (current > 0) showSlide(current - 1);
  });
  nextBtn.addEventListener('click', () => {
    if (current < slides.length - 1) showSlide(current + 1);
  });
  speakBtn.addEventListener('click', speakCurrentSlide);
  stopBtn.addEventListener('click', stopSpeech);

  // Initialize
  showSlide(0);
  startTimer();

  // Preload voices
  if ('speechSynthesis' in window) speechSynthesis.getVoices();
})();

// ============== EXAM VALIDATION ==============
(function() {
  const form = document.getElementById('examForm');
  if (!form) return;
  form.addEventListener('submit', (e) => {
    const total = form.querySelectorAll('.question-card').length;
    const answered = form.querySelectorAll('input[type=radio]:checked').length;
    if (answered < total) {
      if (!confirm(`You answered ${answered} of ${total} questions. Submit anyway?`)) {
        e.preventDefault();
      }
    }
  });
})();