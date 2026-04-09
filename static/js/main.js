/* main.js – AleDocs
   Inicializações globais e utilitários frontend.
*/

$(document).ready(function () {

  // ── Fechar alertas automaticamente após 6 segundos ──────────
  setTimeout(function () {
    document.querySelectorAll('.alert').forEach(function (el) {
      const bsAlert = bootstrap.Alert.getOrCreateInstance(el);
      bsAlert.close();
    });
  }, 6000);

  // ── Tooltip Bootstrap em todo elemento com data-bs-toggle ───
  document.querySelectorAll('[data-bs-toggle="tooltip"]').forEach(function (el) {
    new bootstrap.Tooltip(el);
  });

});
