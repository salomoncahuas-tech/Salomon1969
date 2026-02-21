/**
 * English Assessment Platform - Client-side JavaScript
 */

document.addEventListener('DOMContentLoaded', function() {
    // Auto-dismiss flash messages after 5 seconds
    document.querySelectorAll('.flash').forEach(function(flash) {
        setTimeout(function() {
            flash.style.opacity = '0';
            flash.style.transition = 'opacity 0.3s';
            setTimeout(function() { flash.remove(); }, 300);
        }, 5000);
    });

    // Auto-uppercase access code inputs
    document.querySelectorAll('.input-code, .input-code-large').forEach(function(input) {
        input.addEventListener('input', function() {
            this.value = this.value.toUpperCase();
        });
    });
});
