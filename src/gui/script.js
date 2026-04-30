document.addEventListener('DOMContentLoaded', () => {
    initTheme();
    
    // Timeout para la Splash Screen
    setTimeout(() => {
        const splash = document.getElementById('splash-screen');
        if (splash) splash.classList.add('fade-out');
        loadConfig();
    }, 2000);
});

function initTheme() {
    const savedTheme = localStorage.getItem('theme');
    const systemPrefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
    if (savedTheme === 'dark' || (!savedTheme && systemPrefersDark)) {
        document.body.classList.add('dark-mode');
    }
}

function toggleTheme() {
    const isDark = document.body.classList.toggle('dark-mode');
    localStorage.setItem('theme', isDark ? 'dark' : 'light');
}

function toggleTerminal() {
    const toggle = document.querySelector('.terminal-section .advanced-toggle');
    const content = document.getElementById('terminal-content');
    toggle.classList.toggle('active');
    content.classList.toggle('active');
}

function showView(viewId) {
    document.querySelectorAll('.view').forEach(v => v.classList.remove('active'));
    document.querySelectorAll('.nav-item').forEach(v => v.classList.remove('active'));
    
    document.getElementById(viewId).classList.add('active');
    document.getElementById(`nav-${viewId}`).classList.add('active');
}

function addLog(message) {
    const logContainer = document.getElementById('log');
    const entry = document.createElement('div');
    entry.className = 'log-entry';
    
    const now = new Date();
    const timestamp = now.getHours().toString().padStart(2, '0') + ':' + 
                      now.getMinutes().toString().padStart(2, '0') + ':' + 
                      now.getSeconds().toString().padStart(2, '0');
                      
    entry.innerHTML = `<span style="color: var(--secondary-color)">[${timestamp}]</span> ${message}`;
    logContainer.appendChild(entry);
    logContainer.scrollTop = logContainer.scrollHeight;

    // Actualizar texto de estado dinámicamente basado en el log
    updateStatusFromLog(message);
}

function updateStatusFromLog(msg) {
    const statusMsg = document.getElementById('status-message');
    if (!statusMsg) return;

    if (msg.includes("HDA login")) statusMsg.innerText = "Accediendo al portal HDA...";
    if (msg.includes("Payments tile clicked")) statusMsg.innerText = "Escaneando tickets OneTime Check...";
    if (msg.includes("Processing ticket")) {
        const match = msg.match(/ticket (\d+\/\d+)/);
        statusMsg.innerText = match ? `Analizando datos del ticket ${match[1]}...` : "Extrayendo datos de HDA...";
    }
    if (msg.includes("Starting SAP Validation")) statusMsg.innerText = "Validando información en SAP...";
    if (msg.includes("STARTING AUTOMATIC SUSPENSION")) statusMsg.innerText = "Aplicando suspensiones en HDA...";
    if (msg.includes("Summary email sent")) statusMsg.innerText = "Enviando reportes por correo...";

    if (msg.includes("END PROCESS")) {
        const isSuccess = msg.includes("status=success");
        setFinalStatus(isSuccess, isSuccess ? "Todas las tareas se realizaron con éxito." : "El proceso terminó con advertencias o errores.");
    }
}

function showToast(message, type = 'info') {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    
    let icon = '';
    if (type === 'success') icon = '<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="20 6 9 17 4 12"></polyline></svg>';
    if (type === 'error') icon = '<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2.5"><circle cx="12" cy="12" r="10"></circle><line x1="15" y1="9" x2="9" y2="15"></line><line x1="9" y1="9" x2="15" y2="15"></line></svg>';
    if (type === 'info') icon = '<svg viewBox="0 0 24 24" width="20" height="20" fill="none" stroke="currentColor" stroke-width="2.5"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="16" x2="12" y2="12"></line><line x1="12" y1="8" x2="12.01" y2="8"></line></svg>';

    toast.innerHTML = `
        <div class="toast-content">
            ${icon}
            <span>${message}</span>
        </div>
        <button class="toast-close" onclick="this.parentElement.remove()">&times;</button>
    `;

    container.appendChild(toast);

    // Auto-dismiss after 4 seconds
    setTimeout(() => {
        toast.classList.add('fade-out');
        setTimeout(() => toast.remove(), 300);
    }, 4000);
}

function setFinalStatus(success, message) {
    const iconContainer = document.getElementById('status-icon-container');
    const title = document.getElementById('status-title');
    const msg = document.getElementById('status-message');
    const card = document.getElementById('status-card');

    if (success) {
        iconContainer.innerHTML = '<svg class="success-icon" viewBox="0 0 24 24" width="40" height="40" fill="none" stroke="currentColor" stroke-width="3"><polyline points="20 6 9 17 4 12"></polyline></svg>';
        title.innerText = "¡Proceso Completado!";
        card.style.borderLeftColor = "var(--success-color)";
    } else {
        iconContainer.innerHTML = '<svg class="error-icon" viewBox="0 0 24 24" width="40" height="40" fill="none" stroke="currentColor" stroke-width="3"><circle cx="12" cy="12" r="10"></circle><line x1="15" y1="9" x2="9" y2="15"></line><line x1="9" y1="9" x2="15" y2="15"></line></svg>';
        title.innerText = "Error en el Proceso";
        card.style.borderLeftColor = "var(--danger-color)";
    }
    msg.innerText = message;
}

// API de Comunicación con Python (pywebview)
async function runAutomation() {
    const btn = document.getElementById('runBtn');
    const statusCard = document.getElementById('status-card');
    const iconContainer = document.getElementById('status-icon-container');
    const title = document.getElementById('status-title');
    const msg = document.getElementById('status-message');
    
    btn.disabled = true;
    btn.innerText = 'Ejecutando...';
    
    // Reset Status Card
    statusCard.style.display = 'block';
    statusCard.style.borderLeftColor = "var(--secondary-color)";
    iconContainer.innerHTML = '<div class="spinner"></div>';
    title.innerText = "Procesando...";
    msg.innerText = "Iniciando motores del bot";
    
    addLog("Iniciando proceso principal...");
    
    try {
        const result = await pywebview.api.run_automation();
        if (result.success) {
            // El proceso es asíncrono en Python, el estado final se detectará vía log
            // Pero por si acaso, si el bridge retorna algo inmediato:
        } else {
            setFinalStatus(false, result.error);
        }
    } catch (e) {
        setFinalStatus(false, "Error de comunicación: " + e);
    } finally {
        btn.disabled = false;
        btn.innerText = 'Iniciar Suspensión HDA';
    }
}

async function loadConfig() {
    try {
        const settings = await pywebview.api.get_config();
        
        // HDA
        document.getElementById('hda_url').value = settings.hda_url || '';
        document.getElementById('hda_username').value = settings.hda_username || '';
        document.getElementById('hda_password').value = settings.hda_password || '';
        
        // SAP
        document.getElementById('sap_username_fms').value = settings.sap_username_fms || '';
        document.getElementById('sap_password_fms').value = settings.sap_password_fms || '';
        document.getElementById('sap_username_afs').value = settings.sap_username_afs || '';
        document.getElementById('sap_password_afs').value = settings.sap_password_afs || '';
        
        // SMTP
        document.getElementById('smtp_host').value = settings.smtp_host || '';
        document.getElementById('smtp_port').value = settings.smtp_port || '';
        document.getElementById('smtp_username').value = settings.smtp_username || '';
        document.getElementById('smtp_password').value = settings.smtp_password || '';
        document.getElementById('mail_summary_recipient').value = settings.mail_summary_recipient || '';
        document.getElementById('mail_afs_recipient').value = settings.mail_afs_recipient || '';
        document.getElementById('mail_fms_recipient').value = settings.mail_fms_recipient || '';
        
    } catch (e) {
        console.error("Error cargando configuración:", e);
    }
}

async function saveConfig() {
    const data = {
        hda_url: document.getElementById('hda_url').value,
        hda_username: document.getElementById('hda_username').value,
        hda_password: document.getElementById('hda_password').value,
        sap_username_fms: document.getElementById('sap_username_fms').value,
        sap_password_fms: document.getElementById('sap_password_fms').value,
        sap_username_afs: document.getElementById('sap_username_afs').value,
        sap_password_afs: document.getElementById('sap_password_afs').value,
        smtp_host: document.getElementById('smtp_host').value,
        smtp_port: parseInt(document.getElementById('smtp_port').value) || 0,
        smtp_username: document.getElementById('smtp_username').value,
        smtp_password: document.getElementById('smtp_password').value,
        mail_summary_recipient: document.getElementById('mail_summary_recipient').value,
        mail_afs_recipient: document.getElementById('mail_afs_recipient').value,
        mail_fms_recipient: document.getElementById('mail_fms_recipient').value
    };
    
    try {
        const result = await pywebview.api.save_config(data);
        if (result.success) {
            showToast("Configuración guardada correctamente.", "success");
        } else {
            showToast("Error al guardar: " + result.error, "error");
        }
    } catch (e) {
        showToast("Error de comunicación: " + e, "error");
    }
}

function toggleAdvancedEmail() {
    const toggle = document.querySelector('.advanced-toggle');
    const content = document.getElementById('advanced-email-content');
    
    toggle.classList.toggle('active');
    content.classList.toggle('active');
}
