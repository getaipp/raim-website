/* ═══════════════════════════════════════════════════════════
   RAIM — The Recruiter AI Movement
   Premium Interactions
   ═══════════════════════════════════════════════════════════ */

(function () {
  'use strict';

  // ── Nav ──
  const nav = document.getElementById('nav');
  const navToggle = document.getElementById('nav-toggle');
  const mobileMenu = document.getElementById('mobile-menu');

  window.addEventListener('scroll', () => {
    nav.classList.toggle('scrolled', window.scrollY > 40);
  }, { passive: true });

  navToggle.addEventListener('click', () => {
    mobileMenu.classList.toggle('open');
    document.body.style.overflow = mobileMenu.classList.contains('open') ? 'hidden' : '';
  });

  mobileMenu.querySelectorAll('a').forEach(link => {
    link.addEventListener('click', () => {
      mobileMenu.classList.remove('open');
      document.body.style.overflow = '';
    });
  });


  // ── Hero scroll shrink + desaturate ──
  const heroWrap = document.getElementById('hero-video-wrap');
  if (heroWrap) {
    const hero = document.getElementById('hero');
    const overlay = heroWrap.querySelector('.hero-overlay');
    const content = heroWrap.querySelector('.hero-content');

    function updateHeroScroll() {
      const rect = hero.getBoundingClientRect();
      const heroH = heroWrap.offsetHeight;
      const scrolled = Math.max(0, -rect.top);
      const progress = Math.min(scrolled / (heroH * 0.6), 1);

      // Scale: 1 → 0.92
      const scale = 1 - progress * 0.08;
      // Border radius: 0 → 28px
      const radius = progress * 28;
      // Overlay darkens and desaturates
      const overlayOpacity = 0.3 + progress * 0.5;
      // Content fades out faster
      const contentOpacity = 1 - progress * 2.5;

      heroWrap.style.transform = `scale(${scale})`;
      heroWrap.style.borderRadius = `${radius}px`;
      overlay.style.background = `linear-gradient(180deg, rgba(0,0,0,${overlayOpacity * 0.6}) 0%, rgba(0,0,0,${overlayOpacity}) 100%)`;
      overlay.style.backdropFilter = `saturate(${1 - progress * 0.8}) brightness(${1 - progress * 0.2})`;
      overlay.style.webkitBackdropFilter = `saturate(${1 - progress * 0.8}) brightness(${1 - progress * 0.2})`;
      content.style.opacity = Math.max(0, contentOpacity);
      content.style.transform = `translateY(${progress * -20}px)`;
    }

    window.addEventListener('scroll', updateHeroScroll, { passive: true });
    updateHeroScroll();
  }

  // ── Scroll reveal ──
  const revealObserver = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        entry.target.classList.add('revealed');
        revealObserver.unobserve(entry.target);
      }
    });
  }, { threshold: 0.08, rootMargin: '0px 0px -80px 0px' });

  document.querySelectorAll('[data-reveal], [data-reveal-stagger]').forEach(el => {
    revealObserver.observe(el);
  });

  // ── Counter animation ──
  const counterObserver = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        const el = entry.target;
        const target = parseInt(el.dataset.count, 10);
        if (!target) return;
        const duration = 2000;
        const start = performance.now();

        function animate(now) {
          const elapsed = now - start;
          const progress = Math.min(elapsed / duration, 1);
          // Premium easing — ease-out cubic
          const eased = 1 - Math.pow(1 - progress, 3);
          el.textContent = Math.round(eased * target);
          if (progress < 1) requestAnimationFrame(animate);
        }

        requestAnimationFrame(animate);
        counterObserver.unobserve(el);
      }
    });
  }, { threshold: 0.5 });

  document.querySelectorAll('[data-count]').forEach(el => {
    counterObserver.observe(el);
  });

  // ── Smooth anchor scroll ──
  document.querySelectorAll('a[href^="#"]').forEach(anchor => {
    anchor.addEventListener('click', (e) => {
      const target = document.querySelector(anchor.getAttribute('href'));
      if (target) {
        e.preventDefault();
        target.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
    });
  });

})();
