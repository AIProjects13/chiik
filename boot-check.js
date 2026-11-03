/* ==========================================================================
   Granjas · Vigilante de arranque
   Script clásico (no módulo): se ejecuta incluso cuando app.js no puede cargar.
   Si la app no inicia, reemplaza el "Cargando…" infinito por un mensaje útil.
   app.js marca window.__granjasIniciada = true en cuanto arranca.
   ========================================================================== */
(function () {
    // Anti-clickjacking: la plataforma nunca debe mostrarse dentro de otra página.
    if (window.top !== window.self) {
        document.documentElement.style.display = 'none';
        try { window.top.location = window.self.location; } catch (e) { /* bloqueado por el navegador */ }
        return;
    }

    var loading = document.getElementById('screen-loading');
    var text = document.getElementById('loading-text');
    if (!loading || !text) return;

    function mostrarFallo(mensaje, conReintento) {
        if (window.__granjasIniciada || loading.hidden) return;
        var spinner = loading.querySelector('.spinner');
        if (spinner) spinner.hidden = true;
        text.textContent = mensaje;
        text.className = 'boot-error';
        if (conReintento) {
            var btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn btn-primary';
            btn.textContent = 'Reintentar';
            btn.addEventListener('click', function () { location.reload(); });
            loading.appendChild(btn);
        }
    }

    // Abierta con doble clic (file://): el navegador bloquea la app por seguridad.
    if (location.protocol === 'file:') {
        mostrarFallo('La plataforma no funciona abierta como archivo. Ábrela desde su dirección web.', false);
        return;
    }

    setTimeout(function () {
        mostrarFallo('No se pudo cargar la plataforma. Revisa tu conexión a internet e inténtalo de nuevo.', true);
    }, 15000);
})();
