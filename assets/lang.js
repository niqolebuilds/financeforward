/* Finance Forward — EN/ID reading toggle.
   Persists choice in localStorage so it holds across pages/devices-session.
   Pair this with the inline anti-flash snippet in each page's <head>:
   <script>try{if(localStorage.getItem('ff_lang')==='id')document.documentElement.classList.add('show-id')}catch(e){}</script>
*/
(function () {
  function apply(lang) {
    document.documentElement.classList.toggle('show-id', lang === 'id');
    document.querySelectorAll('[data-lang-btn]').forEach(function (btn) {
      var active = btn.getAttribute('data-lang-btn') === lang;
      btn.classList.toggle('active', active);
      btn.setAttribute('aria-pressed', active ? 'true' : 'false');
    });
  }

  function getSaved() {
    try { return localStorage.getItem('ff_lang') || 'en'; }
    catch (e) { return 'en'; }
  }

  function save(lang) {
    try { localStorage.setItem('ff_lang', lang); } catch (e) {}
  }

  document.addEventListener('DOMContentLoaded', function () {
    apply(getSaved());
  });

  document.addEventListener('click', function (e) {
    var btn = e.target.closest('[data-lang-btn]');
    if (!btn) return;
    var lang = btn.getAttribute('data-lang-btn');
    save(lang);
    apply(lang);
  });
})();
