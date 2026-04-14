/**
 * Діагностична анкета клієнта — Frontend Logic v2.0
 * 
 * Надійна відправка даних до Google Apps Script (GAS) вебхуку.
 * 
 * Ключові функції:
 * - GAS при кросс-доменному POST робить redirect (302), тому
 *   використовуємо mode: 'no-cors' для надійності.
 * - Валідація кожного поля з індивідуальними повідомленнями.
 * - Автозбереження в localStorage (захист від втрати даних).
 * - Авторозмір textarea при вводі.
 * - Розумна Instagram-валідація (@username, URL, username).
 * - Прогрес-бар з нумерованими бейджами.
 * - Захист від повторної відправки (double submit prevention).
 * - Таймаут на fetch (AbortController, 15 секунд).
 */

document.addEventListener('DOMContentLoaded', () => {
    // =========================================================
    // КОНФІГУРАЦІЯ — Вставте ваш URL Google Apps Script сюди
    // =========================================================
    const WEBHOOK_URL = 'https://script.google.com/macros/s/AKfycbz_Bhwb7zXEfMNQ5eN-2rtUncn1AW8xpj8IsawstwZMWdty_hpSEEJbMuXufUTcCk0D/exec';
    // =========================================================

    const STORAGE_KEY = 'brief_form_draft';
    const FETCH_TIMEOUT = 15000; // 15 секунд

    // ─── DOM елементи ──────────────────────────────────
    const form = document.getElementById('briefForm');
    const submitBtn = document.getElementById('submitBtn');
    const btnText = submitBtn.querySelector('.btn-text');
    const btnLoader = submitBtn.querySelector('.btn-loader');
    const successMessage = document.getElementById('successMessage');
    const errorMessage = document.getElementById('errorMessage');
    const retryBtn = document.getElementById('retryBtn');
    const progressFill = document.getElementById('progressFill');
    const progressText = document.getElementById('progressText');
    const progressBar = document.getElementById('progressBar');
    const autosaveIndicator = document.getElementById('autosaveIndicator');

    const REQUIRED_FIELDS = ['q1', 'q2', 'q3', 'q4', 'q5', 'q6', 'q7', 'q8', 'q9', 'q10', 'q11', 'q12'];
    const TOTAL_QUESTIONS = 13;
    let isSubmitting = false;
    let saveTimeout = null;

    // ─── localStorage: Автозбереження ──────────────────
    function saveDraft() {
        const data = {};
        for (let i = 1; i <= TOTAL_QUESTIONS; i++) {
            const field = document.getElementById(`q${i}`);
            if (field && field.value.trim()) {
                data[`q${i}`] = field.value;
            }
        }

        if (Object.keys(data).length > 0) {
            try {
                localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
                showAutosave();
            } catch (e) {
                // localStorage може бути недоступний (private mode)
            }
        }
    }

    function restoreDraft() {
        try {
            const saved = localStorage.getItem(STORAGE_KEY);
            if (!saved) return;

            const data = JSON.parse(saved);
            let restored = false;

            for (const [key, value] of Object.entries(data)) {
                const field = document.getElementById(key);
                if (field && !field.value) {
                    field.value = value;
                    restored = true;
                    // Для textarea — оновити розмір
                    if (field.tagName === 'TEXTAREA') {
                        autoResizeTextarea(field);
                    }
                }
            }

            if (restored) {
                updateProgress();
                updateFilledBadges();
                // Оновити лічильники символів
                document.querySelectorAll('textarea').forEach(ta => {
                    updateCharCount(ta);
                });
            }
        } catch (e) {
            // Ігноруємо помилки парсингу
        }
    }

    function clearDraft() {
        try {
            localStorage.removeItem(STORAGE_KEY);
        } catch (e) {
            // Ігноруємо
        }
    }

    function showAutosave() {
        if (!autosaveIndicator) return;
        autosaveIndicator.textContent = 'Збережено';
        autosaveIndicator.classList.add('visible');
        clearTimeout(saveTimeout);
        saveTimeout = setTimeout(() => {
            autosaveIndicator.classList.remove('visible');
        }, 2000);
    }

    // Дебаунс збереження (500ms після останнього вводу)
    let debounceSaveTimer = null;
    function debounceSave() {
        clearTimeout(debounceSaveTimer);
        debounceSaveTimer = setTimeout(saveDraft, 500);
    }

    // ─── Авторозмір textarea ───────────────────────────
    function autoResizeTextarea(el) {
        // Скидаємо висоту щоб scrollHeight був актуальним
        el.style.height = 'auto';
        const newHeight = Math.min(el.scrollHeight, 400); // max 400px
        el.style.height = newHeight + 'px';
    }

    // ─── Лічильник символів ────────────────────────────
    function updateCharCount(field) {
        const id = field.id;
        const counter = document.querySelector(`.char-count[data-for="${id}"]`);
        if (counter) {
            const len = field.value.length;
            if (len > 0) {
                counter.textContent = `${len} символів`;
                counter.classList.add('active');
            } else {
                counter.textContent = '';
                counter.classList.remove('active');
            }
        }
    }

    // ─── Бейджі заповнених питань ──────────────────────
    function updateFilledBadges() {
        for (let i = 1; i <= TOTAL_QUESTIONS; i++) {
            const field = document.getElementById(`q${i}`);
            const group = field?.closest('.form-group');
            if (field && group) {
                group.classList.toggle('is-filled', field.value.trim().length > 0);
            }
        }
    }

    // ─── Прогрес-бар ─────────────────────────────────────
    function updateProgress() {
        let filled = 0;
        for (let i = 1; i <= TOTAL_QUESTIONS; i++) {
            const field = document.getElementById(`q${i}`);
            if (field && field.value.trim().length > 0) {
                filled++;
            }
        }
        const percent = Math.round((filled / TOTAL_QUESTIONS) * 100);
        progressFill.style.width = `${percent}%`;
        progressText.textContent = `${filled} з ${TOTAL_QUESTIONS} питань`;

        // ARIA
        if (progressBar) {
            progressBar.setAttribute('aria-valuenow', filled);
        }
    }

    // ─── Ініціалізація полів ──────────────────────────────
    for (let i = 1; i <= TOTAL_QUESTIONS; i++) {
        const field = document.getElementById(`q${i}`);
        if (!field) continue;

        // Загальні обробники
        field.addEventListener('input', () => {
            updateProgress();
            updateFilledBadges();
            debounceSave();
        });

        // Textarea: авторозмір + лічильник символів
        if (field.tagName === 'TEXTAREA') {
            field.addEventListener('input', () => {
                autoResizeTextarea(field);
                updateCharCount(field);
            });
        }
    }

    // ─── Автоформат дати народження (ДД.ММ.РРРР) ──────
    const birthdayField = document.getElementById('q2');
    if (birthdayField) {
        birthdayField.addEventListener('input', (e) => {
            let v = birthdayField.value.replace(/[^\d]/g, ''); // тільки цифри
            if (v.length > 8) v = v.slice(0, 8);
            
            // Авто-вставка крапок: DD.MM.YYYY
            let formatted = '';
            for (let i = 0; i < v.length; i++) {
                if (i === 2 || i === 4) formatted += '.';
                formatted += v[i];
            }
            birthdayField.value = formatted;
        });
    }

    // ─── Валідація ────────────────────────────────────────
    const VALIDATION_MESSAGES = {
        q1: 'Будь ласка, вкажіть ваше ПІБ',
        q2: 'Будь ласка, вкажіть дату народження',
        q3: 'Будь ласка, вкажіть Instagram сторінку',
        q4: 'Будь ласка, розкажіть про вашу сферу діяльності',
        q5: 'Будь ласка, вкажіть причину придбання послуги',
        q6: 'Будь ласка, опишіть що стало кінцевою точкою',
        q7: 'Будь ласка, вкажіть яку потребу хочете закрити',
        q8: 'Будь ласка, опишіть вашу мрію або ціль',
        q9: 'Будь ласка, опишіть ваші очікування',
        q10: 'Будь ласка, вкажіть теми які хотіли б вивчити',
        q11: 'Будь ласка, опишіть ваші труднощі',
        q12: 'Будь ласка, розкажіть про минулий досвід'
    };

    /**
     * Розумна валідація Instagram.
     * Приймає: @username, username, https://instagram.com/username,
     *          https://www.instagram.com/username
     * Нормалізує до URL для бекенду.
     */
    function normalizeInstagram(value) {
        let v = value.trim();
        if (!v) return '';

        // Прибираємо зайві пробіли та спецсимволи
        v = v.replace(/\s+/g, '');

        // Якщо це вже URL — повертаємо як є
        if (/^https?:\/\/(www\.)?instagram\.com\/.+/i.test(v)) {
            return v;
        }

        // Прибираємо @ якщо є
        if (v.startsWith('@')) {
            v = v.substring(1);
        }

        // Якщо це схоже на username (лише букви, цифри, крапки, підкреслення)
        if (/^[a-zA-Z0-9._]{1,30}$/.test(v)) {
            return `https://instagram.com/${v}`;
        }

        // Повертаємо як є — можливо це інший формат
        return v;
    }

    function validateForm() {
        let isValid = true;
        let firstError = null;

        // Очищаємо попередні помилки
        document.querySelectorAll('.field-error').forEach(el => {
            el.textContent = '';
            el.classList.remove('visible');
        });
        document.querySelectorAll('.form-group.has-error').forEach(el => {
            el.classList.remove('has-error');
        });

        REQUIRED_FIELDS.forEach(id => {
            const field = document.getElementById(id);
            const errorEl = document.getElementById(`${id}-error`);
            if (field && field.value.trim() === '') {
                isValid = false;
                field.closest('.form-group').classList.add('has-error');
                if (errorEl) {
                    errorEl.textContent = VALIDATION_MESSAGES[id] || 'Це поле обов\'язкове';
                    errorEl.classList.add('visible');
                }
                if (!firstError) firstError = field;
            }
        });

        if (!isValid && firstError) {
            firstError.scrollIntoView({ behavior: 'smooth', block: 'center' });
            // Затримка фокусу щоб скрол встиг відпрацювати
            setTimeout(() => {
                firstError.focus({ preventScroll: true });
            }, 400);
        }

        return isValid;
    }

    // Прибираємо помилку при вводі
    REQUIRED_FIELDS.forEach(id => {
        const field = document.getElementById(id);
        if (field) {
            field.addEventListener('input', () => {
                if (field.value.trim() !== '') {
                    field.closest('.form-group').classList.remove('has-error');
                    const errorEl = document.getElementById(`${id}-error`);
                    if (errorEl) {
                        errorEl.textContent = '';
                        errorEl.classList.remove('visible');
                    }
                }
            });
        }
    });

    // ─── Відправка форми ──────────────────────────────────
    form.addEventListener('submit', async (e) => {
        e.preventDefault();

        // Захист від повторної відправки
        if (isSubmitting) return;

        // Валідація
        if (!validateForm()) return;

        isSubmitting = true;
        setLoadingState(true);
        errorMessage.classList.add('hidden');

        // Збираємо дані
        const data = {};
        for (let i = 1; i <= TOTAL_QUESTIONS; i++) {
            const field = document.getElementById(`q${i}`);
            if (field) {
                let value = field.value.trim();
                // Нормалізуємо Instagram
                if (field.name === 'instagram') {
                    value = normalizeInstagram(value);
                }
                data[field.name] = value;
            }
        }

        try {
            if (WEBHOOK_URL === 'ВАШ_ВЕБХУК_URL_СЮДИ') {
                // Тестовий режим — імітація відправки
                console.warn('⚠️ ВЕБХУК НЕ ПІДКЛЮЧЕНО! Симулюємо відправку для тесту.');
                console.table(data);
                await new Promise(r => setTimeout(r, 2000));
                showSuccess();
                return;
            }

            // AbortController для таймауту
            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), FETCH_TIMEOUT);

            // Відправляємо через no-cors (GAS робить redirect 302,
            // тому ми не можемо прочитати відповідь, але дані будуть збережені)
            await fetch(WEBHOOK_URL, {
                method: 'POST',
                mode: 'no-cors',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify(data),
                redirect: 'follow',
                signal: controller.signal
            });

            clearTimeout(timeoutId);

            // При no-cors ми отримаємо opaque response (status 0),
            // але дані були надіслані успішно
            showSuccess();

        } catch (error) {
            console.error('Помилка відправки:', error);

            if (error.name === 'AbortError') {
                console.error('Таймаут відправки (15 сек)');
            }

            showError();
        }
    });

    // ─── Retry Button ─────────────────────────────────
    if (retryBtn) {
        retryBtn.addEventListener('click', () => {
            isSubmitting = false;
            setLoadingState(false);
            errorMessage.classList.add('hidden');
            // Скролімо до кнопки submit
            submitBtn.scrollIntoView({ behavior: 'smooth', block: 'center' });
        });
    }

    // ─── UI-стани ─────────────────────────────────────
    function setLoadingState(loading) {
        submitBtn.disabled = loading;
        if (loading) {
            btnText.classList.add('hidden');
            btnLoader.classList.remove('hidden');
        } else {
            btnText.classList.remove('hidden');
            btnLoader.classList.add('hidden');
        }
    }

    function showSuccess() {
        form.classList.add('hidden');
        const header = document.getElementById('formHeader');
        if (header) header.classList.add('hidden');
        successMessage.classList.remove('hidden');
        window.scrollTo({ top: 0, behavior: 'smooth' });
        // Очищаємо чернетку після успішної відправки
        clearDraft();
    }

    function showError() {
        isSubmitting = false;
        setLoadingState(false);
        errorMessage.classList.remove('hidden');
        errorMessage.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }

    // ─── Smooth появлення секцій при скролі ──────────────
    const observer = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting) {
                entry.target.classList.add('visible');
                observer.unobserve(entry.target);
            }
        });
    }, { threshold: 0.1, rootMargin: '0px 0px -40px 0px' });

    document.querySelectorAll('.form-section').forEach(section => {
        observer.observe(section);
    });

    // ─── Захист від втрати даних при закритті ─────────────
    window.addEventListener('beforeunload', (e) => {
        if (isSubmitting) return; // Не блокуємо під час відправки

        // Перевіряємо чи є заповнені поля
        let hasFilled = false;
        for (let i = 1; i <= TOTAL_QUESTIONS; i++) {
            const field = document.getElementById(`q${i}`);
            if (field && field.value.trim().length > 0) {
                hasFilled = true;
                break;
            }
        }

        if (hasFilled && !successMessage.classList.contains('hidden') === false) {
            saveDraft(); // Останнє збереження
            e.preventDefault();
            e.returnValue = '';
        }
    });

    // ─── Ініціалізація ────────────────────────────────
    restoreDraft();
    updateProgress();
    updateFilledBadges();
});
