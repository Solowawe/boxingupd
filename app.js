/**
 * Приложение для управления соревнованиями по боксу
 * Backend: Flask + SQLite
 * Frontend: Vanilla JS + Fetch API
 */

// ======================== УВЕДОМЛЕНИЯ (TOAST) ========================

function toast(message, type = 'info', duration = 3500) {
    let container = document.getElementById('toast-container');
    if (!container) {
        container = document.createElement('div');
        container.id = 'toast-container';
        document.body.appendChild(container);
    }
    const el = document.createElement('div');
    el.className = 'toast ' + type;
    el.textContent = message;
    container.appendChild(el);
    requestAnimationFrame(() => el.classList.add('show'));
    setTimeout(() => {
        el.classList.remove('show');
        setTimeout(() => el.remove(), 300);
    }, duration);
}

// ======================== API КЛИЕНТ ========================

const API = {
    base: '/api',

    async request(method, path, body = null) {
        const opts = {
            method,
            headers: { 'Content-Type': 'application/json' },
        };
        if (body !== null) {
            opts.body = JSON.stringify(body);
        }
        const res = await fetch(this.base + path, opts);
        if (!res.ok) {
            const err = await res.json().catch(() => ({ error: res.statusText }));
            throw new Error(err.error || 'Ошибка запроса');
        }
        return res.json();
    },

    // Спортсмены
    getAthletes(params = '') { return this.request('GET', `/athletes${params}`); },
    getAthlete(id) { return this.request('GET', `/athletes/${id}`); },
    getAthletes() { return this.request('GET', '/athletes'); },
    addAthlete(data) { return this.request('POST', '/athletes', data); },
    updateAthlete(id, data) { return this.request('PUT', `/athletes/${id}`, data); },
    deleteAthlete(id) { return this.request('DELETE', `/athletes/${id}`); },

    // Соревнования
    getTournaments() { return this.request('GET', '/tournaments'); },
    getTournament(id) { return this.request('GET', `/tournaments/${id}`); },
    getTournamentCount(id) { return this.request('GET', `/tournaments/${id}/athlete-count`); },
    addTournament(data) { return this.request('POST', '/tournaments', data); },
    updateTournament(id, data) { return this.request('PUT', `/tournaments/${id}`, data); },
    deleteTournament(id) { return this.request('DELETE', `/tournaments/${id}`); },
    getWeightClasses(id) { return this.request('GET', `/tournaments/${id}/weight-classes`); },
    async generateBracket(id, body = null) {
        const res = await fetch(this.base + `/tournaments/${id}/generate`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: body ? JSON.stringify(body) : undefined
        });
        const data = await res.json().catch(() => ({}));
        return { success: res.ok, ...data, status: res.status };
    },
    updateAthletes(tournamentId, athleteIds) { return this.request('PUT', `/tournaments/${tournamentId}/athletes`, { athlete_ids: athleteIds }); },
    addTournamentAthlete(tournamentId, athleteId) { return this.request('POST', `/tournaments/${tournamentId}/athletes`, { athlete_id: athleteId }); },
    removeTournamentAthlete(tournamentId, athleteId) { return this.request('DELETE', `/tournaments/${tournamentId}/athletes/${athleteId}`); },

// Бои
    getFights(tournamentId) { return this.request('GET', `/tournaments/${tournamentId}/fights`); },
    addFight(tournamentId, data) { return this.request('POST', `/tournaments/${tournamentId}/fights`, data); },
    updateFight(fightId, data) { return this.request('PUT', `/fights/${fightId}`, data); },
    
    // Жеребьевка
    getLottery(tournamentId) { return this.request('GET', `/tournaments/${tournamentId}/lottery`); },
    saveLottery(tournamentId, data) { return this.request('POST', `/tournaments/${tournamentId}/lottery`, data); },
    
    // Обновление спортсмена в бою
    updateFightAthlete(tournamentId, fightId, slot, athleteId) {
        return this.request('PUT', `/tournaments/${tournamentId}/fights/${fightId}/athlete/${slot}`, { athlete_id: athleteId });
    },
    
    // Очистка сетки
    clearBracket(tournamentId) {
        return this.request('POST', `/tournaments/${tournamentId}/clear`, {});
    },
    
    // Выбор победителя в бою
    updateFightResult(tournamentId, fightId, winnerId) {
        return this.request('PUT', `/tournaments/${tournamentId}/fights/${fightId}/result`, { winner_id: winnerId });
    },

    // Отмена/сброс результата боя
    resetFightResult(tournamentId, fightId) {
        return this.request('DELETE', `/tournaments/${tournamentId}/fights/${fightId}/result`);
    },

    // Результаты пар сетки (победители из жеребьевки)
    getBracketResults(tournamentId) {
        return this.request('GET', `/tournaments/${tournamentId}/bracket-results`);
    },
    saveBracketResult(tournamentId, weightClass, pairIndex, winnerId, method, day = 1) {
        return this.request('POST', `/tournaments/${tournamentId}/bracket-results`, {
            weight_class: weightClass,
            pair_index: pairIndex,
            winner_id: winnerId,
            method: method || '',
            day: day
        });
    },

    // Очистка результатов пар сетки (всех или одной категории)
    clearBracketResults(tournamentId) {
        return this.request('DELETE', `/tournaments/${tournamentId}/bracket-results`);
    },
    clearWeightClassResults(tournamentId, weightClass) {
        return this.request(
            'DELETE',
            `/tournaments/${tournamentId}/weight-classes/${encodeURIComponent(weightClass)}/results`
        );
    },
};

// ======================== ОСНОВНОЕ ПРИЛОЖЕНИЕ ========================

const App = {
    currentTab: 'tournament',
    currentFightId: null,
    currentTournamentId: null,
    editAthleteId: null,
    athletesCache: [],

    // ==================== УПРАВЛЕНИЕ СОРЕВНОВАНИЕМ ПО УМОЛЧАНИЮ ====================

    saveDefaultTournament(id) {
        if (id) {
            localStorage.setItem('default_tournament_id', id);
        } else {
            localStorage.removeItem('default_tournament_id');
        }
    },

    getDefaultTournamentId() {
        return localStorage.getItem('default_tournament_id');
    },

    // Применяет соревнование по умолчанию ко всем выпадающим спискам
    // (Сетка обрабатывается отдельно в renderBracketSelect)
    async applyDefaultTournament() {
        const defaultId = this.getDefaultTournamentId();
        if (!defaultId) return;

        // Жеребьевка
        const lotterySelect = document.getElementById('lottery-tournament-select');
        if (lotterySelect) {
            const opt = lotterySelect.querySelector(`option[value="${defaultId}"]`);
            if (opt) {
                lotterySelect.value = defaultId;
                // Триггерим событие change для загрузки данных жеребьевки
                lotterySelect.dispatchEvent(new Event('change'));
                // Загружаем результаты жеребьевки из БД
                try {
                    const tournament = await API.getTournament(parseInt(defaultId));
                    await this.performLottery(parseInt(defaultId));
                } catch (err) {
                    console.error('Ошибка загрузки жеребьевки:', err);
                }
            }
        }

        // Результаты
        const resultsSelect = document.getElementById('results-tournament-select');
        if (resultsSelect) {
            const opt = resultsSelect.querySelector(`option[value="${defaultId}"]`);
            if (opt) {
                resultsSelect.value = defaultId;
                resultsSelect.dispatchEvent(new Event('change'));
                try {
                    const tournament = await API.getTournament(parseInt(defaultId));
                    await this.renderResults(tournament);
                } catch (err) {
                    console.error('Ошибка загрузки результатов:', err);
                }
            }
        }

        // Взвешивание
        const weighinSelect = document.getElementById('weighin-tournament-select');
        if (weighinSelect) {
            const opt = weighinSelect.querySelector(`option[value="${defaultId}"]`);
            if (opt) {
                weighinSelect.value = defaultId;
                weighinSelect.dispatchEvent(new Event('change'));
                try {
                    await this.renderWeighin();
                } catch (err) {
                    console.error('Ошибка загрузки взвешивания:', err);
                }
            }
        }

        // Участники — загружаем список для соревнования по умолчанию
        const participantSelect = document.getElementById('participant-tournament-select');
        if (participantSelect) {
            const opt = participantSelect.querySelector(`option[value="${defaultId}"]`);
            if (opt) {
                participantSelect.value = defaultId;
                try {
                    await this.renderParticipants();
                } catch (err) {
                    console.error('Ошибка загрузки участников:', err);
                }
            }
        }
    },

    async checkUpdates() {
        const statusEl = document.getElementById('update-status');
        try {
            const res = await fetch('/api/update/status').then(r => r.json());
            if (!res.configured) {
                if (statusEl) statusEl.textContent = '';
                toast('Обновления не настроены. Укажите GITHUB_OWNER и GITHUB_REPO в server.py (переменные окружения).', 'warning', 6000);
                return;
            }
            if (res.error) {
                if (statusEl) statusEl.textContent = 'Не удалось проверить: ' + res.error;
                toast('Ошибка проверки: ' + res.error, 'error');
                return;
            }
            if (res.up_to_date) {
                if (statusEl) statusEl.textContent = `Актуальная версия ${res.local_version}`;
                toast('Обновлений нет. Текущая версия: ' + (res.local_version || '—'), 'success');
                return;
            }
            if (!confirm(`Доступна новая версия ${res.latest_version} (текущая: ${res.local_version || '—'}). Файлы: ${(res.files || []).join(', ')}. Обновить?`)) return;
            const r = await fetch('/api/update/apply', { method: 'POST' }).then(x => x.json());
            if (r.ok) {
                if (statusEl) statusEl.textContent = `Обновлено до ${r.version}`;
                toast('✅ Приложение успешно обновлено. Обновите страницу (F5). Если обновлялся backend — перезапустите сервер.', 'success', 6000);
            } else {
                if (statusEl) statusEl.textContent = 'Ошибка: ' + r.error;
                toast('Ошибка обновления: ' + r.error, 'error');
            }
        } catch (e) {
            if (statusEl) statusEl.textContent = 'Ошибка: ' + e.message;
            toast('Ошибка: ' + e.message, 'error');
        }
    },

    openInfo(title, body) {
        const m = document.getElementById('info-modal');
        if (!m) return;
        document.getElementById('info-modal-title').textContent = title || 'Справка';
        document.getElementById('info-modal-body').textContent = body || '';
        m.style.display = 'block';
    },
    closeInfo() {
        const m = document.getElementById('info-modal');
        if (m) m.style.display = 'none';
    },
    openTerms() {
        this.openInfo('Пользовательское соглашение',
`ПОЛЬЗОВАТЕЛЬСКОЕ СОГЛАШЕНИЕ

Настоящее соглашение определяет условия использования программы «Соревнования по боксу».

1. ДАННЫЕ ПОЛЬЗОВАТЕЛЯ
1.1. Все данные, вводимые пользователем (спортсмены, соревнования, результаты), хранятся локально на устройстве пользователя и не передаются третьим лицам.

2. ОТВЕТСТВЕННОСТЬ
2.1. Программа предоставляется «как есть». Правообладатель не несёт ответственности за прямой или косвенный ущерб, возникший в результате использования или невозможности использования программы.
2.2. Пользователь самостоятельно проверяет корректность вводимых данных и итоговых результатов.

3. ЗАКЛЮЧИТЕЛЬНЫЕ ПОЛОЖЕНИЯ
3.1. Используя программу, вы подтверждаете согласие с настоящим соглашением.
3.2. По вопросам соглашения обращайтесь к правообладателю.`);
    },

    openLicense() {
        this.openInfo('Лицензионное соглашение',
`ПРОГРАММНОЕ ОБЕСПЕЧЕНИЕ «СОРЕВНОВАНИЯ ПО БОКСУ»

Настоящее лицензионное соглашение (далее — Соглашение) регулирует порядок использования программы для ЭВМ «Соревнования по боксу» (далее — Программа).

1. ПРЕДМЕТ СОГЛАШЕНИЯ
1.1. Правообладатель предоставляет вам неисключительную лицензию на использование Программы в личных (некоммерческих) и профессиональных целях: ведение списка спортсменов, соревнований, жеребьёвки, турнирных сеток и результатов.

2. УСТАНАВЛИВАЕМЫЕ ГРАНИЦЫ ИСПОЛЬЗОВАНИЯ
2.1. Разрешается: устанавливать и запускать Программу на любом количестве устройств; копировать и распространять Программу без изменений; использовать сгенерированные данные в своей деятельности.
2.2. Запрещается: модифицировать, декомпилировать или извлекать исходный код без отдельного разрешения правообладателя; использовать Программу для незаконной деятельности.

3. ДАННЫЕ ПОЛЬЗОВАТЕЛЯ
3.1. Все данные, введённые пользователем (спортсмены, соревнования, результаты), хранятся локально на устройстве пользователя и не передаются третьим лицам.

4. ОГРАНИЧЕНИЕ ОТВЕТСТВЕННОСТИ
4.1. Программа предоставляется «как есть». Правообладатель не несёт ответственности за прямой или косвенный ущерб, возникший в результате использования или невозможности использования Программы.
4.2. Пользователь самостоятельно проверяет корректность вводимых данных и итоговых результатов.

5. ЗАКЛЮЧИТЕЛЬНЫЕ ПОЛОЖЕНИЯ
5.1. Используя Программу, вы подтверждаете согласие с условиями настоящего Соглашения.
5.2. По вопросам Соглашения обращайтесь к правообладателю.`);
    },
async openAbout() {
        let version = '—';
        try {
            const r = await fetch('/update-manifest.json', { cache: 'no-store' });
            if (r.ok) {
                const man = await r.json();
                if (man && man.version) version = man.version;
            }
        } catch (e) { /* игнорируем */ }
        this.openInfo('О приложении',
`СОРЕВНОВАНИЯ ПО БОКСУ

Программа для организации соревнований по боксу.

Возможности:
• Ведение списка спортсменов и весовых категорий
• Импорт участников из Word и Excel (умное сопоставление полей)
• Создание соревнований и расписания по дням
• Жеребьёвка по олимпийской системе
• Турнирная сетка с бронзовыми местами
• Результаты и протоколы, экспорт в Excel и печать

Версия: ${version}
© 2026. Все данные хранятся локально.
По всем вопросам обращаться на почту: liteapps08@ya.ru`);
    },

    async init() {
        this.setupTabs();
        // Закрытие окна «Справка» кликом по фону
        const infoModal = document.getElementById('info-modal');
        if (infoModal) {
            infoModal.addEventListener('click', (e) => {
                if (e.target === infoModal) infoModal.style.display = 'none';
            });
        }
        this.setupExcelImport();
        this.setupWordImport();
        this.setupAthletesImport();
        this.setupExport();
        this.setupTournamentForm();
        this.setupParticipantForm();
        this.setupLottery();
        this.setupFightResultForm();
        this.setupBracketControls();
        this.setupWeighin();
        await this.renderAll();

        // Загружаем список соревнований на активной вкладке "Соревнование"
        await this.renderTournaments();

        // Сначала заполняем селекты соревнований во всех вкладках
        await this.renderBracketSelect();
        await this.renderResultsSelect();

        // Применяем соревнование по умолчанию ко всем разделам
        await this.applyDefaultTournament();

        // Загружаем данные сетки и жеребьевки из БД (после заполнения селектов)
        await this.loadBracketFromDB();
        await this.loadLotteryFromDB();

        // Глобальный обработчик для кнопки "Очистить сетку"
        const clearBtn = document.getElementById('clear-bracket');
        if (clearBtn) {
            clearBtn.addEventListener('click', async () => {
                const select = document.getElementById('bracket-tournament-select');
                const id = parseInt(select.value);
                if (!id) {
                    toast('Выберите соревнование!');
                    return;
                }
                if (!confirm('Сбросить сетку в первоначальный вид? Все результаты боёв будут удалены.')) return;
                
                try {
                    console.log('=== clearBracket click ===');
                    console.trace('Clear bracket stack trace');
                    console.log('Очистка сетки для турнира', id);
                    await API.clearBracket(id);
                    console.log('Сетка очищена');
                    toast('✅ Сетка сброшена в первоначальный вид!');
                    const updatedTournament = await API.getTournament(id);
                    await this.renderBracket(updatedTournament);
                    await this.renderTournaments();
                    // Обновляем результаты, если открыт тот же турнир
                    const rs = document.getElementById('results-tournament-select');
                    if (rs && String(rs.value) === String(id)) {
                        await this.renderResults(updatedTournament);
                    }
                } catch (err) {
                    console.error('Ошибка при очистке сетки:', err);
                    toast('Ошибка: ' + (err.message || 'Неизвестная ошибка'));
                }
            });
        }
    },

    // ==================== АВТОЗАГРУЗКА ДАННЫХ ИЗ БД ====================

    async loadBracketFromDB() {
        console.log('=== loadBracketFromDB START ===');
        const defaultId = this.getDefaultTournamentId();
        console.log('defaultId:', defaultId);
        if (!defaultId) {
            console.log('Нет defaultId — выходим');
            return;
        }

        try {
            console.log('Загружаем турнир', defaultId);
            const tournament = await API.getTournament(parseInt(defaultId));
            console.log('tournament.fights.length:', tournament.fights ? tournament.fights.length : 'undefined');
            console.log('tournament:', tournament);
            
            const bracketSelect = document.getElementById('bracket-tournament-select');
            const resultsSelect = document.getElementById('results-tournament-select');
            
            if (bracketSelect) {
                bracketSelect.value = defaultId;
                console.log('Установили bracketSelect.value =', defaultId);
                await this.loadWeightClasses(parseInt(defaultId));
            }
            
            if (tournament.fights && tournament.fights.length > 0) {
                console.log('Есть бои — рендерим сетку');
                await this.renderBracket(tournament);
            } else {
                console.log('Нет боёв — показываем пустое состояние');
                const container = document.getElementById('bracket-container');
                container.innerHTML = '<div class="empty-state"><p>Сетка не сформирована. Нажмите "Сформировать сетку"</p></div>';
            }
            
            if (resultsSelect) {
                resultsSelect.value = defaultId;
                await this.renderResults(tournament);
            }
        } catch (err) {
            console.error('Ошибка загрузки сетки из БД:', err);
        }
        console.log('=== loadBracketFromDB END ===');
    },

    async loadLotteryFromDB() {
        // Проверяем, есть ли турнир по умолчанию
        const defaultId = this.getDefaultTournamentId();
        if (!defaultId) return;

        try {
            const tournament = await API.getTournament(parseInt(defaultId));
            
            // Если есть участники - отображаем жеребьевку
            if (tournament.athlete_ids && tournament.athlete_ids.length >= 2) {
                const lotterySelect = document.getElementById('lottery-tournament-select');
                if (lotterySelect) {
                    lotterySelect.value = defaultId;
                    // Триггерим событие change для загрузки данных жеребьевки
                    lotterySelect.dispatchEvent(new Event('change'));
                }
            }
        } catch (err) {
            console.error('Ошибка загрузки жеребьевки из БД:', err);
        }
    },

    // ==================== НАВИГАЦИЯ ====================

    setupTabs() {
        document.querySelectorAll('.nav-btn').forEach(btn => {
            btn.addEventListener('click', async () => {
                const tab = btn.dataset.tab;
                await this.switchTab(tab);
            });
        });
    },

    async switchTab(tab) {
        this.currentTab = tab;
        document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
        document.querySelector(`.nav-btn[data-tab="${tab}"]`).classList.add('active');
        document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
        document.getElementById(`tab-${tab}`).classList.add('active');

        if (tab === 'tournament') await this.renderTournaments();
        if (tab === 'bracket') {
            // Всегда вызываем renderBracketSelect для установки обработчиков
            await this.renderBracketSelect();
            // Автоматически формируем сетку при переключении на вкладку
            const select = document.getElementById('bracket-tournament-select');
            if (select && select.value) {
                const tournament = await API.getTournament(parseInt(select.value));
                await this.renderBracket(tournament);
            }
        }
        if (tab === 'results') {
            await this.renderResultsSelect();
        }
        if (tab === 'athletes') {
            await this.setupAthleteTournamentSelect();
            await this.renderAthletes();
            await this.renderParticipants();
        }
        if (tab === 'weighin') await this.renderWeighin();
    },

    // ==================== СПОРТСМЕНЫ ====================

    setupParticipantForm() {
        const select = document.getElementById('participant-tournament-select');
        const athleteSelect = document.getElementById('participant-athlete-select');
        const form = document.getElementById('participant-form');

        const updateParticipantControls = () => {
            const hasTournament = !!select.value;
            if (athleteSelect) athleteSelect.disabled = !hasTournament;
            const addBtn = document.getElementById('add-participant-btn');
            if (addBtn) addBtn.disabled = !hasTournament;
        };
        updateParticipantControls();

        // Кнопка "Добавить вручную"
        const addManualBtn = document.getElementById('add-manual-athlete-btn');
        const manualForm = document.getElementById('manual-athlete-form');
        const manualInputForm = document.getElementById('manual-athlete-input-form');
        const cancelManualBtn = document.getElementById('cancel-manual-athlete');

        if (addManualBtn) {
            addManualBtn.addEventListener('click', () => {
                manualForm.style.display = manualForm.style.display === 'none' ? 'block' : 'none';
                if (manualForm.style.display === 'block') {
                    document.getElementById('manual-name').focus();
                }
            });
        }

        if (cancelManualBtn) {
            cancelManualBtn.addEventListener('click', () => {
                manualForm.style.display = 'none';
                manualInputForm.reset();
                document.getElementById('manual-athlete-id').value = '';
                document.getElementById('manual-form-title').textContent = '➕ Новый спортсмен';
                document.getElementById('manual-submit-btn').textContent = '💾 Добавить';
            });
        }

        // Кнопки поиска участников
        const searchInput = document.getElementById('participants-search');
        const searchBtn = document.getElementById('participants-search-btn');
        const clearBtn = document.getElementById('participants-clear-btn');

        if (searchBtn) {
            searchBtn.addEventListener('click', () => this.renderParticipants());
        }
        if (searchInput) {
            searchInput.addEventListener('keypress', (e) => {
                if (e.key === 'Enter') this.renderParticipants();
            });
        }
        if (clearBtn) {
            clearBtn.addEventListener('click', () => {
                searchInput.value = '';
                this.renderParticipants();
            });
        }

        if (manualInputForm) {
            manualInputForm.addEventListener('submit', async (e) => {
                e.preventDefault();
                const athleteId = document.getElementById('manual-athlete-id').value;
                const tournamentId = manualInputForm.dataset.tournamentId;

                const name = document.getElementById('manual-name').value.trim();
                const birthYear = document.getElementById('manual-birth-year').value;
                const weight = document.getElementById('manual-weight').value;
                const gender = document.getElementById('manual-gender').value;
                const rank = document.getElementById('manual-rank').value.trim();
                const region = document.getElementById('manual-region').value.trim();
                const coach = document.getElementById('manual-coach').value.trim();

                if (!name || !birthYear || !weight) {
                    toast('Заполните обязательные поля: Ф.И.О., Год рождения, Вес!');
                    return;
                }

                try {
                    let athlete;
                    if (athleteId) {
                        // Редактирование существующего
                        athlete = await API.updateAthlete(parseInt(athleteId), {
                            name,
                            birth_year: parseInt(birthYear),
                            weight: parseFloat(weight),
                            gender,
                            rank,
                            region,
                            coach
                        });
                        toast('✅ Данные сохранены!');
                    } else {
                        // Добавление нового
                        athlete = await API.addAthlete({
                            name,
                            birth_year: parseInt(birthYear),
                            weight: parseFloat(weight),
                            gender,
                            rank,
                            region,
                            coach
                        });

                        // Добавляем в выпадающий список
                        const opt = document.createElement('option');
                        opt.value = athlete.id;
                        opt.textContent = `${athlete.name} (${athlete.weight} кг, ${athlete.gender})`;
                        athleteSelect.appendChild(opt);
                        athleteSelect.value = athlete.id;

                        // Автоматически привязываем к турниру по умолчанию
                        const defaultId = this.getDefaultTournamentId();
                        if (defaultId) {
                            try {
                                await API.addTournamentAthlete(parseInt(defaultId), athlete.id);
                            } catch (err) {
                                console.error('Ошибка привязки к турниру:', err);
                            }
                        }
                    }

                    // Закрываем форму и обновляем список
                    const manualForm = document.getElementById('manual-athlete-form');
                    manualForm.style.display = 'none';
                    manualInputForm.reset();
                    document.getElementById('manual-athlete-id').value = '';
                    document.getElementById('manual-form-title').textContent = '➕ Новый спортсмен';
                    document.getElementById('manual-submit-btn').textContent = '💾 Добавить';

                    // Загружаем спортсменов для обновления выпадающего списка
                    const athletes = await API.getAthletes();
                    athleteSelect.innerHTML = '<option value="">-- Выберите спортсмена --</option>';
                    athletes.forEach(a => {
                        const opt = document.createElement('option');
                        opt.value = a.id;
                        opt.textContent = `${a.name} (${a.weight} кг, ${a.gender})`;
                        athleteSelect.appendChild(opt);
                    });

                    await this.renderParticipants();
                    await this.renderTournaments();
                    toast('✅ Спортсмен добавлен и привязан к турниру!');
                } catch (err) {
                    toast('Ошибка: ' + err.message);
                }
            });
        }

        // Загружаем соревнования (переиспользуемо — вызывается и после создания турнира)
        this.loadParticipantSelect = async () => {
            try {
                const tournaments = await API.getTournaments();
                select.innerHTML = '<option value="">-- Выберите соревнование --</option>';
                tournaments.forEach(t => {
                    const opt = document.createElement('option');
                    opt.value = t.id;
                    opt.textContent = `${t.name} (${t.date_from || t.date_to || '—'}) [${t.athlete_count || 0} уч.]`;
                    select.appendChild(opt);
                });

                // Загружаем сохраненное соревнование по умолчанию
                const defaultId = this.getDefaultTournamentId();
                if (defaultId) {
                    const opt = select.querySelector(`option[value="${defaultId}"]`);
                    if (opt) opt.selected = true;
                }
                updateParticipantControls();
            } catch (err) {
                console.error('Ошибка загрузки соревнований:', err);
            }
        };

        this.loadParticipantSelect();
        select.addEventListener('change', async () => {
            updateParticipantControls();
            await this.renderParticipants();
        });

        // Загружаем участников при инициализации
        select.addEventListener('DOMContentLoaded', () => this.renderParticipants());

        // Загружаем спортсменов для выбора
        const loadAthletes = async () => {
            try {
                const athletes = await API.getAthletes();
                athleteSelect.innerHTML = '<option value="">-- Выберите спортсмена --</option>';
                athletes.forEach(a => {
                    const opt = document.createElement('option');
                    opt.value = a.id;
                    opt.textContent = `${a.name} (${a.weight} кг, ${a.gender})`;
                    athleteSelect.appendChild(opt);
                });
            } catch (err) {
                console.error('Ошибка загрузки спортсменов:', err);
            }
        };

        loadAthletes();

        // Добавление участника
        form.addEventListener('submit', async (e) => {
            e.preventDefault();
            const tournamentId = parseInt(select.value);
            const addSelect = document.getElementById('participant-add-select');
            const athleteId = parseInt(addSelect ? addSelect.value : '');

            if (!tournamentId || !athleteId) {
                toast('Выберите соревнование и спортсмена!');
                return;
            }

            try {
                await API.addTournamentAthlete(tournamentId, athleteId);
                if (addSelect) addSelect.value = '';
                await this.renderParticipants();
                if (this.loadParticipantSelect) await this.loadParticipantSelect(); // Обновляем селект турнира
                await this.renderTournaments(); // Обновляем счётчики на карточках
            } catch (err) {
                toast('Ошибка: ' + err.message);
            }
        });
    },

    async renderParticipants() {
        try {
            const tournamentId = document.getElementById('participant-tournament-select')?.value;
            const tbody = document.getElementById('participants-list');
            const count = document.getElementById('participants-count');
            const searchInput = document.getElementById('participants-search');
            const searchTerm = searchInput ? searchInput.value.toLowerCase().trim() : '';

            if (!tournamentId) {
                tbody.innerHTML = '';
                count.textContent = 'Участников: 0';
                return;
            }

            const tournament = await API.getTournament(parseInt(tournamentId));
            const athletes = await API.getAthletes();
            const tournamentAthleteIds = tournament.athlete_ids || [];
            let participants = athletes.filter(a => tournamentAthleteIds.includes(a.id));

            // Применяем поиск
            if (searchTerm) {
                participants = participants.filter(a =>
                    a.name.toLowerCase().includes(searchTerm) ||
                    (a.region || '').toLowerCase().includes(searchTerm) ||
                    (a.coach || '').toLowerCase().includes(searchTerm) ||
                    (a.rank || '').toLowerCase().includes(searchTerm)
                );
            }

            tbody.innerHTML = '';
            count.textContent = `Участников: ${participants.length}`;
            this._participantsCache = participants;

            // Список «Участники турнира» — строго по записям в БД (tournament_athletes)
            const selParts = document.getElementById('participant-athlete-select');
            if (selParts) {
                if (participants.length === 0) {
                    selParts.innerHTML = '<option value="">-- Участников пока нет --</option>';
                } else {
                    selParts.innerHTML = participants.map(p =>
                        `<option value="${p.id}">${this.escapeHtml(p.name)} (${p.weight} кг, ${p.gender})</option>`).join('');
                }
            }
            // Список «Добавить спортсмена» — только НЕ состоящие в турнире
            const selAdd = document.getElementById('participant-add-select');
            if (selAdd) {
                const partIds = new Set(tournamentAthleteIds);
                const available = athletes.filter(a => !partIds.has(a.id));
                if (available.length === 0) {
                    selAdd.innerHTML = '<option value="">-- Все спортсмены уже в турнире --</option>';
                } else {
                    selAdd.innerHTML = '<option value="">-- Добавить спортсмена (не в турнире) --</option>' +
                        available.map(a => `<option value="${a.id}">${this.escapeHtml(a.name)} (${a.weight} кг, ${a.gender})</option>`).join('');
                }
            }

            if (participants.length === 0) {
                tbody.innerHTML = '<tr><td colspan="9" class="empty-state">Нет участников. Добавьте спортсменов через форму выше</td></tr>';
                return;
            }

            participants.forEach((a, i) => {
                const tr = document.createElement('tr');
                tr.style.cursor = 'pointer';
                tr.title = 'Двойной клик для редактирования';
                tr.innerHTML = `
                    <td>${i + 1}</td>
                    <td>${this.escapeHtml(a.name)}</td>
                    <td>${a.weight}</td>
                    <td>${a.gender}</td>
                    <td>${a.birth_year}</td>
                    <td>${this.escapeHtml(a.rank || '-')}</td>
                    <td>${this.escapeHtml(a.region || '-')}</td>
                    <td>${this.escapeHtml(a.coach || '-')}</td>
                    <td>
                        <button class="btn-secondary" onclick="App.editParticipant(${tournamentId}, ${a.id})" title="Редактировать данные спортсмена">✏️</button>
                        <button class="btn-danger" onclick="App.removeParticipant(${tournamentId}, ${a.id})" title="Убрать из этого турнира">🗑</button>
                        <button class="btn-danger" onclick="App.deleteAthleteFromDb(${a.id})" title="Удалить из базы полностью" style="margin-left:4px;">✖</button>
                    </td>
                `;
                tr.addEventListener('dblclick', (e) => {
                    if (e.target.tagName === 'BUTTON') return;
                    this.editAthleteFromRow(tournamentId, a);
                });
                tbody.appendChild(tr);
            });
        } catch (err) {
            console.error('Ошибка загрузки участников:', err);
            // Чистим таблицу, чтобы не оставались «фантомные» строки от старого рендера
            tbody.innerHTML = '';
            if (count) count.textContent = 'Участников: 0';
        }
    },

    async removeParticipant(tournamentId, athleteId) {
        if (!confirm('Удалить участника из соревнования?')) return;
        try {
            await API.removeTournamentAthlete(tournamentId, athleteId);
            await this.renderParticipants();
        } catch (err) {
            toast('Ошибка: ' + err.message);
        }
    },

    // Открытие формы редактирования по id участника (через кэш текущего списка)
    editParticipant(tournamentId, id) {
        const a = (this._participantsCache || []).find(x => x.id === id);
        if (!a) { toast('Спортсмен не найден', 'error'); return; }
        this.editAthleteFromRow(tournamentId, a);
    },

    editAthleteFromRow(tournamentId, athlete) {
        const manualForm = document.getElementById('manual-athlete-form');
        const manualInputForm = document.getElementById('manual-athlete-input-form');
        const manualFormTitle = document.getElementById('manual-form-title');
        const manualSubmitBtn = document.getElementById('manual-submit-btn');

        document.getElementById('manual-athlete-id').value = athlete.id;
        document.getElementById('manual-name').value = athlete.name;
        document.getElementById('manual-birth-year').value = athlete.birth_year;
        document.getElementById('manual-weight').value = athlete.weight;
        document.getElementById('manual-gender').value = athlete.gender;
        document.getElementById('manual-rank').value = athlete.rank || '';
        document.getElementById('manual-region').value = athlete.region || '';
        document.getElementById('manual-coach').value = athlete.coach || '';

        manualFormTitle.textContent = '✏️ Редактирование: ' + athlete.name;
        manualSubmitBtn.textContent = '💾 Сохранить';
        manualForm.style.display = 'block';
        manualInputForm.dataset.tournamentId = tournamentId;
        manualInputForm.dataset.athleteId = athlete.id;
        document.getElementById('manual-name').focus();
    },

    async exportTournamentToExcel(tournamentId) {
        try {
            const tournament = await API.getTournament(tournamentId);
            const athletes = await API.getAthletes();
            const tournamentAthleteIds = tournament.athlete_ids || [];
            const participants = athletes.filter(a => tournamentAthleteIds.includes(a.id));

            if (participants.length === 0) {
                toast('В соревновании нет участников!');
                return;
            }

            // Создаем CSV (можно открыть в Excel)
            let csv = '№;Ф.И.О.;Вес (кг);Пол;Год рождения;Разряд;Регион;Тренер\n';
            participants.forEach((a, i) => {
                csv += `${i + 1};${a.name};${a.weight};${a.gender};${a.birth_year};${a.rank || ''};${a.region || ''};${a.coach || ''}\n`;
            });

            const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
            const url = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            link.download = `tournament_${tournament.name.replace(/\s+/g, '_')}_participants.csv`;
            link.click();
            URL.revokeObjectURL(url);
        } catch (err) {
            toast('Ошибка экспорта: ' + err.message);
        }
    },

    // Экспорт протокола результатов соревнования в Excel (3 листа: участники, бои, призёры)
    exportProtocolToExcel(tournamentId) {
        try {
            API.getTournament(tournamentId).then(async (tournament) => {
                const athletes = await API.getAthletes();
                const idMap = {};
                athletes.forEach(a => idMap[a.id] = a);
                const participants = (tournament.athlete_ids || []).map(id => idMap[id]).filter(Boolean);

                // Лист 1 — участники
                const sheetMain = [['№', 'Ф.И.О.', 'Вес (кг)', 'Пол', 'Год рождения', 'Разряд', 'Регион', 'Тренер']];
                participants.forEach((a, i) => {
                    sheetMain.push([i + 1, a.name, a.weight, a.gender, a.birth_year, a.rank || '', a.region || '', a.coach || '']);
                });

                // Лист 2 — бои и результаты
                const fights = (tournament.fights || []).slice().sort((a, b) =>
                    (a.weight_class === b.weight_class ? a.round - b.round : a.weight_class < b.weight_class ? -1 : 1) || a.id - b.id);
                const sheetFights = [['Весовая категория', 'Раунд', 'Боец 1', 'Боец 2', 'Победитель', 'Способ', 'Завершён']];
                fights.forEach(f => {
                    const a1 = f.athlete1_id ? idMap[f.athlete1_id] : null;
                    const a2 = f.athlete2_id ? idMap[f.athlete2_id] : null;
                    const win = f.winner_id ? idMap[f.winner_id] : null;
                    sheetFights.push([
                        f.weight_class || '',
                        f.round || '',
                        a1 ? a1.name : '(бай)',
                        a2 ? a2.name : '(бай)',
                        win ? win.name : '',
                        f.method || '',
                        f.completed ? 'Да' : 'Нет'
                    ]);
                });

                // Лист 3 — призёры (протокол): победитель последнего раунда каждой категории
                const sheetProtocol = [['Весовая категория', 'Место', 'Ф.И.О.', 'Регион', 'Тренер']];
                const byClass = {};
                fights.forEach(f => {
                    if (!byClass[f.weight_class]) byClass[f.weight_class] = [];
                    if (f.winner_id) byClass[f.weight_class].push({ v: f.winner_id, r: f.round || 1 });
                });
                Object.keys(byClass).sort().forEach(wc => {
                    // Победитель категории — победитель боя с самым большим раундом
                    const finals = byClass[wc].sort((x, y) => y.r - x.r);
                    const goldId = finals.length ? finals[0].v : null;
                    if (goldId && idMap[goldId]) {
                        sheetProtocol.push([wc, 1, idMap[goldId].name, idMap[goldId].region || '', idMap[goldId].coach || '']);
                    }
                });

                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sheetMain), 'Участники');
                XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sheetFights), 'Бои');
                XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sheetProtocol), 'Призёры');
                XLSX.writeFile(wb, `protocol_${(tournament.name || 'tournament').replace(/\s+/g, '_')}.xlsx`);
                toast('✅ Протокол сохранён', 'success');
            });
        } catch (err) {
            toast('Ошибка экспорта протокола: ' + err.message);
        }
    },

    // Печать результатов текущего соревнования
    async printResults() {
        this.switchTab('results');
        const select = document.getElementById('results-tournament-select');
        const tid = parseInt(select.value);
        if (!tid) {
            toast('Выберите соревнование!');
            return;
        }
        try {
            const tournament = await API.getTournament(tid);
            await this.renderResults(tournament);
            setTimeout(() => window.print(), 300);
        } catch (err) {
            toast('Ошибка печати: ' + err.message, 'error');
        }
    },

    setupSearchAndFilter() {
        // Больше не используется
    },

    setupAthleteForm() {
        // Больше не используется
    },

    setupAthleteForm() {
        document.getElementById('athlete-form').addEventListener('submit', async (e) => {
            e.preventDefault();
            const name = document.getElementById('athlete-name').value.trim();
            const birthYear = document.getElementById('athlete-birth-year').value;
            const weight = document.getElementById('athlete-weight').value;
            const gender = document.getElementById('athlete-gender').value;
            const rank = document.getElementById('athlete-rank').value.trim();
            const region = document.getElementById('athlete-region').value.trim();
            const coach = document.getElementById('athlete-coach').value.trim();

            if (!name || !birthYear || !weight) {
                toast('Заполните обязательные поля: Ф.И.О., Год рождения, Вес!');
                return;
            }

            try {
                if (this.editAthleteId) {
                    await API.updateAthlete(this.editAthleteId, { name, birth_year: birthYear, weight, gender, rank, region, coach });
                    this.editAthleteId = null;
                } else {
                    await API.addAthlete({ name, birth_year: birthYear, weight, gender, rank, region, coach });
                }
                document.getElementById('athlete-form').reset();
                await this.renderAthletes();
            } catch (err) {
                toast('Ошибка: ' + err.message);
            }
        });
    },

    setupBracketControls() {
        // Инициализируем элементы управления сеткой
        const select = document.getElementById('bracket-tournament-select');
        if (select) {
            // Если уже есть опции, загружаем данные
            if (select.options.length > 1) {
                this.renderBracketSelect();
            }
        }
    },

    async renderAthletes() {
        try {
            const search = document.getElementById('athlete-search')?.value || '';
            const gender = document.getElementById('athlete-filter-gender')?.value || '';
            const tournamentId = document.getElementById('athlete-tournament-select')?.value;

            let athletes = await API.getAthletes();

            // Фильтруем по соревнованию если выбрано
            if (tournamentId) {
                const tournament = await API.getTournament(parseInt(tournamentId));
                const tournamentAthleteIds = tournament.athlete_ids || [];
                athletes = athletes.filter(a => tournamentAthleteIds.includes(a.id));
            }

            // Применяем поиск и фильтр
            if (search) {
                athletes = athletes.filter(a => a.name.toLowerCase().includes(search.toLowerCase()));
            }
            if (gender) {
                athletes = athletes.filter(a => a.gender === gender);
            }
            this.athletesCache = athletes;

            const tbody = document.getElementById('athletes-list');
            const count = document.getElementById('athletes-count');

            tbody.innerHTML = '';
            count.textContent = `Всего спортсменов: ${athletes.length}`;

            if (athletes.length === 0) {
                tbody.innerHTML = '<tr><td colspan="10" class="empty-state">Спортсмены не найдены</td></tr>';
                return;
            }

            athletes.forEach((a, i) => {
                const tr = document.createElement('tr');
                tr.dataset.athleteId = a.id;
                tr.innerHTML = `
                    <td>${i + 1}</td>
                    <td class="editable" data-field="name">${this.escapeHtml(a.name)}</td>
                    <td class="editable" data-field="birth_year">${a.birth_year}</td>
                    <td class="editable" data-field="weight">${a.weight}</td>
                    <td class="editable" data-field="gender">${a.gender}</td>
                    <td class="editable" data-field="rank">${this.escapeHtml(a.rank || '-')}</td>
                    <td class="editable" data-field="region">${this.escapeHtml(a.region || '-')}</td>
                    <td class="editable" data-field="coach">${this.escapeHtml(a.coach || '-')}</td>
                    <td class="action-cell">
                        <button class="btn-secondary btn-edit" onclick="App.toggleEditRow(${a.id})" style="margin-right:4px;">✏️</button>
                        <button class="btn-save" onclick="App.saveRow(${a.id})" style="display:none;margin-right:4px;">💾</button>
                        <button class="btn-cancel" onclick="App.cancelEditRow(${a.id})" style="display:none;margin-right:4px;">✖</button>
                        <button class="btn-danger" onclick="App.deleteAthlete(${a.id})">🗑</button>
                    </td>
                `;
                tbody.appendChild(tr);
            });
        } catch (err) {
            console.error('Ошибка загрузки спортсменов:', err);
        }
    },

    toggleEditRow(id) {
        const tr = document.querySelector(`tr[data-athlete-id="${id}"]`);
        if (!tr) return;

        const athlete = this.athletesCache.find(a => a.id === id);
        if (!athlete) return;

        const cells = tr.querySelectorAll('td.editable');
        cells.forEach(cell => {
            const field = cell.dataset.field;
            if (field === 'gender') {
                const val = cell.textContent.trim();
                cell.innerHTML = `<select class="inline-edit" data-field="gender">
                    <option value="мужской" ${val === 'мужской' ? 'selected' : ''}>Мужской</option>
                    <option value="женский" ${val === 'женский' ? 'selected' : ''}>Женский</option>
                </select>`;
            } else {
                const val = cell.textContent.trim();
                const placeholder = val === '-' ? '' : val;
                cell.innerHTML = `<input class="inline-edit" data-field="${field}" value="${this.escapeHtml(placeholder)}">`;
            }
        });

        // Переключаем кнопки
        tr.querySelector('.btn-edit').style.display = 'none';
        tr.querySelector('.btn-save').style.display = 'inline-block';
        tr.querySelector('.btn-cancel').style.display = 'inline-block';
    },

    cancelEditRow(id) {
        const tr = document.querySelector(`tr[data-athlete-id="${id}"]`);
        if (!tr) return;

        tr.querySelector('.btn-edit').style.display = 'inline-block';
        tr.querySelector('.btn-save').style.display = 'none';
        tr.querySelector('.btn-cancel').style.display = 'none';

        // Просто перерисовываем — значения вернутся из кеша
        this.renderAthletes();
    },

    getRowData(id) {
        const tr = document.querySelector(`tr[data-athlete-id="${id}"]`);
        if (!tr) return null;

        const data = { id };
        const inputs = tr.querySelectorAll('.inline-edit');
        inputs.forEach(inp => {
            data[inp.dataset.field] = inp.value.trim();
        });
        return data;
    },

    async saveRow(id) {
        const data = this.getRowData(id);
        if (!data) return;

        if (!data.name || !data.birth_year || !data.weight) {
            toast('Заполните обязательные поля: Ф.И.О., Год рождения, Вес!');
            return;
        }

        try {
            await API.updateAthlete(id, {
                name: data.name,
                birth_year: parseInt(data.birth_year),
                weight: parseFloat(data.weight),
                gender: data.gender,
                rank: data.rank || '',
                region: data.region || '',
                coach: data.coach || ''
            });
            await this.renderAthletes();
        } catch (err) {
            toast('Ошибка сохранения: ' + err.message);
        }
    },

    async deleteAthlete(id) {
        if (!confirm('Удалить спортсмена?')) return;
        try {
            await API.deleteAthlete(id);
            await this.renderAthletes();
        } catch (err) {
            toast('Ошибка: ' + err.message);
        }
    },

    // Полное удаление спортсмена из базы (отвязка от всех турниров)
    async deleteAthleteFromDb(id) {
        if (!confirm('Удалить спортсмена из базы полностью? Он будет отвязан от всех турниров, а его данные удалены.')) return;
        try {
            await API.deleteAthlete(id);
            await this.renderParticipants();
            await this.renderTournaments();
            if (this.loadParticipantSelect) await this.loadParticipantSelect();
        } catch (err) {
            toast('Ошибка: ' + err.message, 'error');
        }
    },

    // Удалить ВСЕХ спортсменов из базы
    async deleteAllAthletes() {
        if (!confirm('Удалить ВСЕХ спортсменов из базы? Это действие необратимо — будут удалены все бойцы и их связи с турнирами.')) return;
        try {
            const r = await fetch('/api/athletes/clear', { method: 'DELETE' });
            if (!r.ok) throw new Error('Не удалось очистить');
            toast('✅ Все спортсмены удалены', 'success');
            await this.renderParticipants();
            await this.renderTournaments();
            if (this.loadParticipantSelect) await this.loadParticipantSelect();
        } catch (e) {
            toast('Ошибка: ' + e.message, 'error');
        }
    },

    setupExcelImport() {
        const btn = document.getElementById('import-excel-btn');
        if (!btn) return;
        
        btn.addEventListener('click', async () => {
            const fileInput = document.getElementById('excel-import');
            const file = fileInput.files[0];
            if (!file) {
                toast('Выберите Excel-файл!');
                return;
            }

            const reader = new FileReader();
            reader.onload = async (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    const rows = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

                    // Получаем соревнование по умолчанию
                    const defaultTournamentId = this.getDefaultTournamentId();
                    let tournamentId = null;
                    let tournamentName = '';
                    
                    if (defaultTournamentId) {
                        const tournament = await API.getTournament(defaultTournamentId);
                        tournamentId = tournament.id;
                        tournamentName = tournament.name;
                    }

                    let imported = 0;
                    let errors = 0;
                    let skipped = 0;
                    const addedAthleteIds = [];

                    // Сначала загружаем всех существующих спортсменов
                    const existingAthletes = await API.getAthletes();

                    for (const [index, row] of rows.entries()) {
                        if (index === 0) continue;
                        if (!row || row.length < 2) continue;

                        const name = String(row[0] || '').trim();
                        const birthYear = parseInt(row[1]);
                        const weight = parseFloat(String(row[2] || '0').replace(',', '.'));
                        let gender = 'мужской';
                        let rank = '';
                        let region = '';
                        let coach = '';

                        if (row[3]) {
                            const g = String(row[3]).trim().toLowerCase();
                            if (g.includes('жен') || g === 'ж' || g === 'female') {
                                gender = 'женский';
                            }
                        }
                        if (row[4]) rank = String(row[4]).trim();
                        if (row[5]) region = String(row[5]).trim();
                        if (row[6]) coach = String(row[6]).trim();

                        // Мягкая валидация — принимаем любой разумный год рождения
                        if (name && !isNaN(birthYear) && birthYear > 1900 && birthYear < 2020) {
                            // Проверяем, есть ли уже спортсмен с таким именем
                            const existing = existingAthletes.find(
                                a => a.name.toLowerCase() === name.toLowerCase()
                            );

                            if (existing) {
                                // Спортсмен уже есть — используем его ID
                                addedAthleteIds.push(existing.id);
                                skipped++;
                            } else {
                                try {
                                    const athlete = await API.addAthlete({ name, birth_year: birthYear, weight, gender, rank, region, coach });
                                    imported++;
                                    existingAthletes.push(athlete); // Обновляем кеш
                                    
                                    // Привязываем к соревнованию по умолчанию
                                    if (tournamentId && athlete && athlete.id) {
                                        addedAthleteIds.push(athlete.id);
                                    }
                                } catch (err) {
                                    console.error('Ошибка добавления спортсмена:', err);
                                    errors++;
                                }
                            }
                        } else {
                            errors++;
                        }
                    }

                    // Привязываем всех импортированных спортсменов к соревнованию по умолчанию
                    if (tournamentId && addedAthleteIds.length > 0) {
                        for (const athleteId of addedAthleteIds) {
                            try {
                                await API.addTournamentAthlete(tournamentId, athleteId);
                            } catch (err) {
                                console.error('Ошибка привязки к соревнованию:', err);
                            }
                        }
                    }

                    await this.renderParticipants();
                    await this.renderTournaments(); // Обновляем счётчики
                    fileInput.value = '';
                    
                    const parts = [`✅ Импорт завершен!`];
                    if (imported > 0) parts.push(`Добавлено новых: ${imported}`);
                    if (skipped > 0) parts.push(`Уже были в базе: ${skipped}`);
                    if (errors > 0) parts.push(`Пропущено: ${errors}`);
                    
                    if (tournamentId) {
                        parts.push(`Все участники привязаны к соревнованию "${tournamentName}"!`);
                    } else {
                        parts.push(`Соревнование по умолчанию не выбрано — спортсмены добавлены, но не привязаны`);
                    }
                    
                    toast(parts.join('\n'));
                } catch (err) {
                    toast('Ошибка при чтении файла: ' + err.message);
                }
            };
            reader.readAsArrayBuffer(file);
        });
    },

    setupWordImport() {
        const btn = document.getElementById('import-word-btn');
        if (!btn) return;
        btn.addEventListener('click', async () => {
            const fileInput = document.getElementById('word-import');
            const file = fileInput.files[0];
            if (!file) { toast('Выберите Word-файл!'); return; }
            const gender = document.getElementById('word-import-gender')?.value || 'мужской';
            const wantReplace = document.getElementById('word-replace')?.checked;
            if (wantReplace && !confirm('Заменить текущий список спортсменов? ВСЕ существующие бойцы (включая текущих) будут удалены, останется только импортированный состав.')) {
                return;
            }
            try {
                const fd = new FormData();
                fd.append('file', file);
                fd.append('gender', gender);
                if (wantReplace) fd.append('replace', '1');
                const res = await fetch('/api/import/docx', { method: 'POST', body: fd });
                const data = await res.json();
                if (!res.ok || data.error) {
                    toast('❌ ' + (data.error || 'Ошибка импорта'), 'error');
                    return;
                }
                // Привязываем новых бойцов к соревнованию по умолчанию
                const defaultId = this.getDefaultTournamentId();
                const ids = data.ids || [];
                if (defaultId && ids.length > 0) {
                    for (const aid of ids) {
                        try { await API.addTournamentAthlete(parseInt(defaultId), aid); }
                        catch (e) { /* ignore */ }
                    }
                }
                await this.renderParticipants();
                await this.renderTournaments();
                fileInput.value = '';
                const msg = `✅ Импортировано: ${data.imported}` + (data.errors ? ` | Ошибок: ${data.errors}` : '');
                toast(msg, 'success', 5000);
                if (defaultId && ids.length > 0) toast(`Привязано к соревнованию по умолчанию: ${ids.length}`, 'info', 4000);
            } catch (e) {
                toast('Ошибка: ' + e.message, 'error');
            }
        });
    },

setupAthletesImport() {
        const btn = document.getElementById('import-athletes-btn');
        if (!btn) return;
        btn.addEventListener('click', async () => {
            const fileInput = document.getElementById('athletes-import');
            const file = fileInput.files[0];
            if (!file) { toast('Выберите файл (Word или Excel)!'); return; }

            const modeRadio = document.querySelector('input[name="import-mode"]:checked');
            const replaceMode = !!(modeRadio && modeRadio.value === 'replace');
            if (replaceMode && !confirm('⚠️ Режим «Заменить список»: при импорте ВСЯ текущая информация о спортсменах будет удалена и заменена загружаемым списком. Продолжить?')) return;

            const ext = (file.name.split('.').pop() || '').toLowerCase();
            const gender = document.getElementById('athletes-import-gender')?.value || 'мужской';
            const statusEl = document.getElementById('import-status');
            const runBtn = document.getElementById('import-athletes-btn');
            if (statusEl) statusEl.textContent = '⏳ Загрузка…';
            if (runBtn) runBtn.disabled = true;
            // Турнир для привязки: по умолчанию или первый существующий
            let targetTid = null;
            const defId = this.getDefaultTournamentId();
            if (defId) { targetTid = parseInt(defId); }
            else {
                try {
                    const tlist = await API.getTournaments();
                    if (tlist && tlist.length) targetTid = tlist[0].id;
                } catch (e) { /* ignore */ }
            }
            try {
                let importedCount = 0, errorCount = 0, skippedCount = 0, ids = [];
                if (ext === 'docx') {
                    const fd = new FormData();
                    fd.append('file', file);
                    fd.append('gender', gender);
                    if (replaceMode) fd.append('replace', '1'); // в режиме «дополнить» старый список не трогаем
                    const res = await fetch('/api/import/docx', { method: 'POST', body: fd });
                    const data = await res.json();
                    if (!res.ok || data.error) {
                        if (statusEl) statusEl.innerHTML = '❌ Ошибка загрузки: ' + (data.error || 'Ошибка импорта');
                        toast('Ошибка загрузки: ' + (data.error || 'Ошибка импорта'), 'error');
                        if (runBtn) runBtn.disabled = false;
                        return;
                    }
                    importedCount = data.imported; errorCount = data.errors || 0; skippedCount = data.skipped || 0; ids = data.ids || [];
                } else {
                    // Excel (.xls/.xlsx) — при «замене» очищаем список, затем импортируем
                    if (replaceMode) {
                        const clearRes = await fetch('/api/athletes/clear', { method: 'DELETE' });
                        if (!clearRes.ok) throw new Error('Не удалось очистить список');
                    }
                    const r = await this._importExcelRows(file, targetTid);
                    importedCount = r.imported; errorCount = r.errors || 0; skippedCount = r.skipped || 0;
                    ids = r.ids || [];
                }
                // Привязываем новых бойцов к турниру
                if (targetTid && ids.length > 0) {
                    for (const aid of ids) {
                        try { await API.addTournamentAthlete(targetTid, aid); } catch (e) {}
                    }
                }
                if (this.loadParticipantSelect) await this.loadParticipantSelect();
                await this.renderParticipants();
                await this.renderTournaments();
                fileInput.value = '';
                const detail = (replaceMode ? 'Заменён список: ' : 'Дополнено: ') + `${importedCount}`
                    + (skippedCount ? ` (уже были: ${skippedCount})` : '')
                    + (errorCount ? ` | Ошибок: ${errorCount}` : '');
                if (statusEl) statusEl.innerHTML = '✅ Все спортсмены успешно загружены. ' + detail;
                let msg = '✅ Все спортсмены успешно загружены. ' + detail;
                if (targetTid && ids.length > 0) msg += `\nПривязано к турниру #${targetTid}: ${ids.length}`;
                else if (!targetTid) msg += `\nСоревнований нет — турнир не создан, привязка не выполнена`;
                toast(msg, 'success', 6000);
                if (runBtn) runBtn.disabled = false;
            } catch (e) {
                if (statusEl) statusEl.innerHTML = '❌ Ошибка загрузки: ' + e.message;
                toast('Ошибка загрузки: ' + e.message, 'error');
                if (runBtn) runBtn.disabled = false;
            }
        });
    },

    // Читает Excel-файл и добавляет бойцов; список уже очищен вызывающей стороной
    async _importExcelRows(file, tournamentId) {
        const reader = new FileReader();
        return new Promise((resolve, reject) => {
            reader.onerror = () => reject(new Error('Не удалось прочитать файл'));
            reader.onload = async (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const sheet = workbook.Sheets[workbook.SheetNames[0]];
                    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

                    let imported = 0, errors = 0, skipped = 0;
                    const addedIds = [];
                    const existing = await API.getAthletes();

                    // --- Умное сопоставление колонок по заголовкам ---
                    const norm = s => String(s || '').toLowerCase().trim();
                    function pickCol(headerRow, ...keys) {
                        for (let i = 0; i < headerRow.length; i++) {
                            const t = norm(headerRow[i]);
                            if (t && keys.some(k => t.includes(k))) return i;
                        }
                        return -1;
                    }
                    let headerIdx = -1;
                    for (let i = 0; i < rows.length; i++) {
                        const joined = (rows[i] || []).map(norm).join(' ');
                        if (/фио|ф\.и\.о|имя|фамилия|вес|год|тренер/.test(joined)) { headerIdx = i; break; }
                    }
                    let ci; // поле -> индекс колонки
                    if (headerIdx >= 0) {
                        const h = rows[headerIdx] || [];
                        ci = {
                            name:   pickCol(h, 'фио', 'ф.и.о', 'имя', 'фамилия', 'спортсмен'),
                            birth:  pickCol(h, 'год рождения', 'год рожд', 'дата рождения', 'год'),
                            weight: pickCol(h, 'вес'),
                            gender: pickCol(h, 'пол'),
                            rank:   pickCol(h, 'разряд', 'квалификация'),
                            region: pickCol(h, 'регион', 'субъект', 'город'),
                            coach:  pickCol(h, 'тренер', 'физрук')
                        };
                    } else {
                        ci = { name:0, birth:1, weight:2, gender:3, rank:4, region:5, coach:6 };
                        headerIdx = -1;
                    }
                    const dataStart = headerIdx >= 0 ? headerIdx + 1 : 0;
                    const cell = (row, c) => (c >= 0 && row[c] != null) ? String(row[c]).trim() : '';
                    const yearOf = s => { const m = String(s || '').match(/(19|20)\d{2}/); return m ? parseInt(m[0]) : NaN; };
                    const weightOf = s => parseFloat(String(s || '').replace(',', '.').replace(/[^\d.]/g, ''));

                    for (let index = dataStart; index < rows.length; index++) {
                        const row = rows[index];
                        if (!row) continue;
                        const name = cell(row, ci.name);
                        if (!name) continue;
                        const birthYear = yearOf(cell(row, ci.birth));
                        const weight = weightOf(cell(row, ci.weight)) || 0;
                        let gender = 'мужской';
                        const gv = norm(cell(row, ci.gender));
                        if (gv.includes('жен') || gv === 'ж' || gv === 'female') gender = 'женский';
                        const rank = cell(row, ci.rank);
                        const region = cell(row, ci.region);
                        const coach = cell(row, ci.coach);
                        if (!isNaN(birthYear) && birthYear > 1900 && birthYear < 2020) {
                            const dup = existing.find(a => a.name.toLowerCase() === name.toLowerCase());
                            if (dup) {
                                addedIds.push(dup.id);
                                skipped++;
                            } else {
                                try {
                                    const athlete = await API.addAthlete({ name, birth_year: birthYear, weight, gender, rank, region, coach });
                                    imported++;
                                    existing.push(athlete);
                                    if (athlete && athlete.id) addedIds.push(athlete.id);
                                } catch (err) { console.error('Ошибка добавления:', err); errors++; }
                            }
                        } else {
                            errors++;
                        }
                    }
                    // Привязкой занимается вызывающий код (нужен одинаковый флоу для Word и Excel)
                    resolve({ imported, errors, ids: addedIds, skipped });
                } catch (err) {
                    reject(err);
                }
            };
            reader.readAsArrayBuffer(file);
        });
    },

setupExport() {
        // Этот метод больше не используется
    },

    // ==================== ВЗВЕШИВАНИЕ ====================

    setupWeighin() {
        const searchInput = document.getElementById('weighin-search');
        const genderFilter = document.getElementById('weighin-gender-filter');
        const select = document.getElementById('weighin-tournament-select');

        if (searchInput) {
            searchInput.addEventListener('input', () => this.renderWeighin());
        }
        if (genderFilter) {
            genderFilter.addEventListener('change', () => this.renderWeighin());
        }
        if (!select) {
            console.warn('setupWeighin: элемент weighin-tournament-select не найден');
            return;
        }
        const loadTournaments = async () => {
            try {
                const tournaments = await API.getTournaments();
                select.innerHTML = '<option value="">-- Все соревнования --</option>';
                tournaments.forEach(t => {
                    const opt = document.createElement('option');
                    opt.value = t.id;
                    opt.textContent = `${t.name} (${this.fmtDate(t.date_from || t.date_to) || '—'})`;
                    select.appendChild(opt);
                });

                // Загружаем сохраненное соревнование по умолчанию
                const defaultId = this.getDefaultTournamentId();
                if (defaultId) {
                    const opt = select.querySelector(`option[value="${defaultId}"]`);
                    if (opt) opt.selected = true;
                }
            } catch (err) {
                console.error('Ошибка загрузки соревнований:', err);
            }
        };

        loadTournaments();

        select.addEventListener('change', () => this.renderWeighin());
    },

    async renderWeighin() {
        try {
            const search = document.getElementById('weighin-search')?.value || '';
            const gender = document.getElementById('weighin-gender-filter')?.value || '';
            const tournamentId = document.getElementById('weighin-tournament-select')?.value;

            let athletes = await API.getAthletes();

            // Фильтруем по соревнованию если выбрано
            if (tournamentId) {
                const tournament = await API.getTournament(parseInt(tournamentId));
                const tournamentAthleteIds = tournament.athlete_ids || [];
                athletes = athletes.filter(a => tournamentAthleteIds.includes(a.id));
            }

            // Применяем поиск и фильтр
            if (search) {
                athletes = athletes.filter(a => a.name.toLowerCase().includes(search.toLowerCase()));
            }
            if (gender) {
                athletes = athletes.filter(a => a.gender === gender);
            }

            const tbody = document.getElementById('weighin-list');
            const count = document.getElementById('weighin-count');

            tbody.innerHTML = '';
            count.textContent = `Спортсменов: ${athletes.length}`;

            if (athletes.length === 0) {
                tbody.innerHTML = '<tr><td colspan="7" class="empty-state">Спортсмены не найдены</td></tr>';
                return;
            }

            athletes.forEach((a, i) => {
                const tr = document.createElement('tr');
                tr.dataset.athleteId = a.id;
                tr.innerHTML = `
                    <td>${i + 1}</td>
                    <td>${this.escapeHtml(a.name)}</td>
                    <td>${a.birth_year}</td>
                    <td>${a.gender}</td>
                    <td>${this.escapeHtml(a.region || '-')}</td>
                    <td><input type="number" class="weighin-input" data-athlete-id="${a.id}" value="${a.weight}" step="0.1" min="30" max="200"></td>
                    <td><button class="btn-save" onclick="App.saveWeighin(${a.id})">💾</button></td>
                `;
                tbody.appendChild(tr);
            });
        } catch (err) {
            console.error('Ошибка загрузки взвешивания:', err);
        }
    },

    async saveWeighin(id) {
        const input = document.querySelector(`.weighin-input[data-athlete-id="${id}"]`);
        if (!input) return;

        const newWeight = parseFloat(input.value);
        if (isNaN(newWeight) || newWeight < 30 || newWeight > 200) {
            toast('Введите корректный вес (30–200 кг)');
            input.focus();
            return;
        }

        try {
            await API.updateAthlete(id, { weight: newWeight });
            input.style.borderColor = '#2ecc71';
            setTimeout(() => { input.style.borderColor = ''; }, 1500);

            // Получаем ID текущего соревнования
            const tournamentId = document.getElementById('weighin-tournament-select')?.value;
            if (tournamentId) {
                const tid = parseInt(tournamentId);

                // Перезагружаем список спортсменов на вкладке взвешивания
                await this.renderWeighin();

                // Пересоздаём жеребьевку
                try {
                    await this.performLottery(tid);
                } catch (err) {
                    console.error('Ошибка обновления жеребьевки:', err);
                }

                // Пересоздаём сетку через API
                const bracketSelect = document.getElementById('bracket-tournament-select');
                if (bracketSelect && bracketSelect.value == tid) {
                    try {
                        const classesList = document.getElementById('weight-classes-list');
                        const checked = classesList ? 
                            [...classesList.querySelectorAll('input:checked')].map(c => c.value) : 
                            [];
                        const body = { weight_classes: checked };
                        const result = await API.generateBracket(tid, body);
                        if (result.success) {
                            // Перезагружаем турнир с новыми боями
                            const tournament = await API.getTournament(tid);
                            await this.renderBracket(tournament);
                        }
                    } catch (err) {
                        console.error('Ошибка обновления сетки:', err);
                    }
                }
            }
        } catch (err) {
            toast('Ошибка сохранения веса: ' + err.message);
        }
    },

    // ==================== СОРЕВНОВАНИЯ ====================

    setupTournamentForm() {
        document.getElementById('tournament-form').addEventListener('submit', async (e) => {
            e.preventDefault();
            const name = document.getElementById('tournament-name').value.trim();
            const dateFrom = document.getElementById('tournament-date-from').value;
            const dateTo = document.getElementById('tournament-date-to').value;
            const days = parseInt(document.getElementById('tournament-days')?.value) || 1; // счётчик дней закомментирован — по умолчанию 1

            if (!name || !dateFrom) {
                toast('Заполните название и дату!');
                return;
            }

            try {
                const result = await API.addTournament({ name, date_from: dateFrom, date_to: dateTo || dateFrom, days });
                document.getElementById('tournament-form').reset();
                // Новое соревнование становится по умолчанию
                this.saveDefaultTournament(result.id);
                await this.renderTournaments();
                // Обновляем селект турниров во вкладке «Участники»
                if (this.loadParticipantSelect) await this.loadParticipantSelect();
                // Показываем сообщение об успешном создании
                toast('✅ Соревнование создано! Добавь спортсменов через кнопку "👥 Участники"');
            } catch (err) {
                toast('Ошибка: ' + err.message);
            }
        });
    },

    async renderTournaments() {
        const container = document.getElementById('tournaments-container');
        try {
            const tournaments = await API.getTournaments();
            if (!container) return;

            if (tournaments.length === 0) {
                container.innerHTML = '<div class="empty-state"><p>Нет соревнований</p></div>';
                return;
            }

            // Очищаем контейнер перед добавлением
            container.innerHTML = '';

            // Загружаем количество участников для каждого соревнования
            const defaultId = this.getDefaultTournamentId();
            
            for (const t of tournaments) {
                try {
                    const tournament = await API.getTournament(t.id);
                    t.athlete_count = (tournament.athlete_ids || []).length;
                } catch (err) {
                    console.error('Ошибка загрузки участников соревнования:', err);
                }
            }

            tournaments.forEach(t => {
                const dateRange = t.date_from
                            ? `${this.fmtDate(t.date_from)}${t.date_to ? ' — ' + this.fmtDate(t.date_to) : ''}`
                            : '';
                const isDefault = defaultId && parseInt(defaultId) === t.id;
                const card = document.createElement('div');
                card.className = 'tournament-card';
                if (isDefault) card.style.borderLeft = '4px solid #ffd700';
                card.innerHTML = `
                    <div>
                        <div class="name">${this.escapeHtml(t.name)}${isDefault ? ' <span style="color:#ffd700;font-size:0.8em;">(по умолчанию)</span>' : ''}</div>
                        <div class="date">📅 ${dateRange} | Участников: ${t.athlete_count || 0} | Боёв: ${t.completed_fights || 0}/${t.total_fights || 0}</div>
                    </div>
                    <div>
                        <button class="btn-secondary" onclick="App.setDefaultTournament(${t.id})">${isDefault ? '⭐ Уже по умолчанию' : '⬇ По умолчанию'}</button>
                        <button class="btn-secondary" onclick="App.manageParticipants(${t.id})">👥 Участники</button>
                        <button class="btn-secondary" onclick="App.exportTournamentToExcel(${t.id})">📊 Excel</button>
                        <button class="btn-secondary" onclick="App.exportProtocolToExcel(${t.id})">📤 Протокол</button>
                        <button class="btn-secondary" style="color: #ff6b6b;" onclick="App.clearParticipants(${t.id})">🗑 Очистить участников</button>
                        <button class="btn-secondary" onclick="App.deleteTournament(${t.id})">❌ Удалить соревнование</button>
                    </div>
                `;
                container.appendChild(card);
            });
            console.log('renderTournaments: соревнования добавлены в DOM');
        } catch (err) {
            console.error('Ошибка загрузки соревнований:', err);
        }
    },

    setDefaultTournament(id) {
        this.saveDefaultTournament(id);
        toast('Соревнование установлено по умолчанию! Оно будет автоматически выбрано во всех разделах.');
        this.renderTournaments();
    },

    async deleteTournament(id) {
        if (!confirm('Удалить соревнование? Все данные о боях будут потеряны.')) return;
        try {
            await API.deleteTournament(id);
            await this.renderTournaments();
        } catch (err) {
            toast('Ошибка: ' + err.message);
        }
    },

    async clearParticipants(tournamentId) {
        if (!confirm('Удалить ВСЕХ участников из этого соревнования?')) return;
        try {
            const tournament = await API.getTournament(tournamentId);
            const athleteIds = tournament.athlete_ids || [];
            for (const athleteId of athleteIds) {
                try {
                    await API.removeTournamentAthlete(tournamentId, athleteId);
                } catch (err) {
                    console.error('Ошибка удаления участника:', err);
                }
            }
            toast('✅ Все участники удалены из соревнования');
            await this.renderTournaments();
            await this.renderParticipants();
        } catch (err) {
            toast('Ошибка: ' + err.message);
        }
    },

    async manageParticipants(tournamentId) {
        // Получаем данные соревнования и всех спортсменов
        const [tournament, athletes] = await Promise.all([
            API.getTournament(tournamentId),
            API.getAthletes()
        ]);
        const selected = new Set(tournament.athlete_ids || []);

        const listHtml = athletes.map(a => `
            <label class="weight-class-check">
                <input type="checkbox" value="${a.id}" ${selected.has(a.id) ? 'checked' : ''}>
                ${this.escapeHtml(a.name)} (${a.weight} кг, ${a.gender})
            </label>
        `).join('');

        // Простое модальное окно через confirm (заменю на prompt-style)
        const modal = document.createElement('div');
        modal.className = 'modal';
        modal.style.display = 'flex';
        modal.innerHTML = `
            <div class="modal-content" style="max-width:600px;">
                <h2>Участники соревнования</h2>
                <p class="info">Отметьте спортсменов, участвующих в соревновании</p>
                <div style="display:flex; flex-wrap:wrap; gap:4px; max-height:400px; overflow-y:auto; padding:10px 0;">
                    ${listHtml}
                </div>
                <div style="display:flex; gap:10px; margin-top:15px; justify-content:flex-end;">
                    <button id="mng-save" class="btn-primary">💾 Сохранить</button>
                    <button id="mng-cancel" class="btn-secondary">Отмена</button>
                </div>
            </div>
        `;
        document.body.appendChild(modal);

        modal.querySelector('#mng-save').addEventListener('click', async () => {
            const ids = [...modal.querySelectorAll('input:checked')].map(c => parseInt(c.value));
            try {
                await API.updateAthletes(tournamentId, ids);
                modal.remove();
                await this.renderTournaments();
                toast('Участники обновлены');
            } catch (err) {
                toast('Ошибка: ' + err.message);
            }
        });
        modal.querySelector('#mng-cancel').addEventListener('click', () => modal.remove());
        modal.addEventListener('click', (e) => { if (e.target === modal) modal.remove(); });
    },

    // ==================== ЖЕРЕБЬЕВКА ====================

    setupLottery() {
        const select = document.getElementById('lottery-tournament-select');
        const container = document.getElementById('lottery-container');
        const results = document.getElementById('lottery-results');

        const loadTournaments = async () => {
            try {
                const tournaments = await API.getTournaments();
                select.innerHTML = '<option value="">-- Выберите соревнование --</option>';
                tournaments.forEach(t => {
                    const opt = document.createElement('option');
                    opt.value = t.id;
                    opt.textContent = `${t.name} (${this.fmtDate(t.date_from || t.date_to) || '—'})`;
                    select.appendChild(opt);
                });

                // Загружаем сохраненное соревнование по умолчанию
                const defaultId = this.getDefaultTournamentId();
                if (defaultId) {
                    const opt = select.querySelector(`option[value="${defaultId}"]`);
                    if (opt) opt.selected = true;
                }
            } catch (err) {
                console.error('Ошибка загрузки соревнований:', err);
            }
        };

        loadTournaments();

        document.getElementById('generate-lottery').addEventListener('click', async () => {
            const id = parseInt(select.value);
            if (!id) {
                toast('Выберите соревнование!');
                return;
            }
            this.saveDefaultTournament(id);
            try {
                await this.performLottery(id);
            } catch (err) {
                toast('Ошибка жеребьевки: ' + err.message);
            }
        });

        container.innerHTML = '<div class="empty-state"><p>Выберите соревнование и нажмите "Провести жеребьевку"</p></div>';
    },

    // ==================== УМНЫЙ ПОДБОР СОПЕРНИКОВ ДЛЯ ЖЕРЕБЬЁВКИ ====================

    _scorePair(a1, a2) {
        // Вес — главный критерий (1 кг = 1000 очков)
        // Возраст — вторичный (1 год = 1 очко)
        const wDiff = Math.abs(a1.weight - a2.weight);
        const aDiff = Math.abs(a1.birth_year - a2.birth_year);
        return wDiff * 1000 + aDiff * 1;
    },

    _pairAndPlaceLottery(athletes) {
        // Олимпийская система (seeding): сортируем по весу → по возрасту (сильнейшие = нижние номера),
        // затем расставляем по стандартной турнирной сетке, чтобы 1-й и 2-й сеяные
        // встречались только в финале, а сильнейшие не пересекались в ранних раундах.
        athletes.sort((a, b) => a.weight - b.weight || a.birth_year - b.birth_year);

        const n = athletes.length;
        if (n < 1) return [];
        // Один спортсмен без соперника — попадёт в жеребьёвку и станет победителем автоматом
        if (n === 1) return [athletes[0]];

        function seedOrder(m) {
            // m — степень двойки; стандартный порядок расстановки сеянных (1 и 2 в разных половинах)
            if (m <= 2) return [1, 2].slice(0, m);
            const half = seedOrder(m / 2);
            const order = [];
            for (let i = 0; i < half.length; i++) {
                order.push(half[i]);        // верхняя половина
                order.push(m + 1 - half[i]); // зеркальная нижняя
            }
            return order;
        }

        const size = 1 << Math.ceil(Math.log2(n));       // ближайшая сверху степень двойки
        const order = seedOrder(size);             // полный порядок сидений (длина = size)
        // Номер сидения > n — свободный слот (бай), ставим null
        const slots = order.map(seed => seed <= n ? athletes[seed - 1] : null);
        return slots;  // массив длины size, null = бай
    },

    async performLottery(tournamentId) {
        const container = document.getElementById('lottery-container');
        const results = document.getElementById('lottery-results');

        // Получаем данные соревнования и спортсменов
        const [tournament, athletes] = await Promise.all([
            API.getTournament(tournamentId),
            API.getAthletes()
        ]);

        // Фильтруем спортсменов соревнования
        const tournamentAthletes = athletes.filter(a => (tournament.athlete_ids || []).includes(a.id));

        if (tournamentAthletes.length === 0) {
            container.innerHTML = `
                <div class="empty-state">
                    <p>⚠️ В соревновании нет участников!</p>
                    <p>Добавьте спортсменов через кнопку <strong>"👥 Участники"</strong> в разделе "Соревнования"</p>
                </div>
            `;
            return;
        }

        if (tournamentAthletes.length < 2) {
            container.innerHTML = '<div class="empty-state"><p>Нужно минимум 2 спортсмена для жеребьевки</p></div>';
            return;
        }

        // Группируем по весовым категориям (категория = точный вес спортсмена)
        const weightGroups = {};
        tournamentAthletes.forEach(a => {
            if (!a.weight) return;
            const w = Number(a.weight);
            const wStr = Number.isInteger(w) ? String(w) : String(w);
            const key = `${a.gender} ${wStr} кг`;
            weightGroups[key] = weightGroups[key] || [];
            weightGroups[key].push(a);
        });

        // Умный подбор соперников по весу и возрасту (олимпийская система)
        const lotteryResults = {};
        Object.entries(weightGroups).forEach(([weightClass, group]) => {
            const ordered = this._pairAndPlaceLottery(group); // массив слотов, null = бай
            let slot = 1;
            ordered.forEach(ath => {
                if (ath === null) return; // бай — в жеребьёвку не попадает
                lotteryResults[weightClass] = lotteryResults[weightClass] || [];
                lotteryResults[weightClass].push({
                    athlete_id: ath.id,
                    slot: slot++,
                    day: 1,
                    name: ath.name,
                    weight: ath.weight,
                    gender: ath.gender,
                    birth_year: ath.birth_year,
                    region: ath.region,
                    coach: ath.coach
                });
            });
        });

        // Сохраняем в БД
        try {
            await API.saveLottery(tournamentId, lotteryResults);
        } catch (err) {
            console.error('Ошибка сохранения жеребьевки:', err);
        }

        // Рендерим результаты
        container.innerHTML = `
            <div class="info" style="margin:15px 0;">
                <strong>${tournament.name}</strong> — ${tournamentAthletes.length} спортсменов распределено по ${Object.keys(lotteryResults).length} категориям
                <button class="btn-secondary" onclick="App.refreshLottery(${tournamentId})" style="margin-left:10px;">🔄 Перемешать</button>
            </div>
        `;

        // Количество дней из разницы дат турнира (для выбора дня в каждом слоте)
        let daysCount = 1;
        if (tournament.date_from && tournament.date_to) {
            const diff = (new Date(tournament.date_to) - new Date(tournament.date_from)) / 86400000;
            if (isFinite(diff) && diff >= 0) daysCount = Math.floor(diff) + 1;
        }
        const dayOptions = Array.from({ length: daysCount }, (_, i) => i + 1)
            .map(d => `<option value="${d}">День ${d}</option>`).join('');

        let html = '';
        Object.entries(lotteryResults).forEach(([weightClass, athletes]) => {
            html += `<div class="lottery-category">
                <h3 class="weight-title">${this.escapeHtml(weightClass)}</h3>
                <table class="lottery-table">
                    <thead>
                        <tr>
                            <th>Слот</th>
                            <th>День</th>
                            <th>Ф.И.О.</th>
                            <th>Вес</th>
                            <th>Действия</th>
                        </tr>
                    </thead>
                    <tbody>`;
            athletes.forEach(a => {
                const daySelect = `<select class="lottery-day-select" data-slot="${a.slot}">${dayOptions}</select>`;
                html += `
                        <tr data-athlete-id="${a.athlete_id}" data-weight-class="${weightClass}">
                            <td><strong>${a.slot}</strong></td>
                            <td>${daySelect}</td>
                            <td>${this.escapeHtml(a.name)}</td>
                            <td>${a.weight} кг</td>
                            <td><button class="btn-edit-fighter" onclick="App.editLotteryAthlete(${tournamentId}, this, '${this.escapeHtml(weightClass)}', ${a.slot}, ${a.athlete_id})" title="Заменить спортсмена">✏️</button></td>
                        </tr>`;
            });
            html += `
                    </tbody>
                </table>
            </div>`;
        });

        // Кнопка сохранения жеребьёвки (после ручной правки соперников/дней)
        html += '<div style="text-align:center; margin-top:20px;">';
        html += '<button class="btn-primary" onclick="App.saveLotteryFromUI()" style="font-size:1.1em; padding:12px 30px;">💾 Сохранить жеребьевку</button>';
        html += '</div>';

        results.innerHTML = html;
    },

    async refreshLottery(tournamentId) {
        if (!confirm('Перемешать жеребьевку? Текущие настройки будут потеряны.')) return;
        
        const tournament = await API.getTournament(tournamentId);
        const athletes = await API.getAthletes();
        const tournamentAthletes = athletes.filter(a => (tournament.athlete_ids || []).includes(a.id));
        
        const weightGroups = {};
        tournamentAthletes.forEach(a => {
            if (!a.weight) return;
            const w = Number(a.weight);
            const wStr = Number.isInteger(w) ? String(w) : String(w);
            const key = `${a.gender} ${wStr} кг`;
            weightGroups[key] = weightGroups[key] || [];
            weightGroups[key].push(a);
        });

        // Умный подбор соперников по весу и возрасту (олимпийская система)
        const lotteryResults = {};
        Object.entries(weightGroups).forEach(([weightClass, group]) => {
            const ordered = this._pairAndPlaceLottery(group); // массив слотов, null = бай
            let slot = 1;
            ordered.forEach(ath => {
                if (ath === null) return; // бай — в жеребьёвку не попадает
                lotteryResults[weightClass] = lotteryResults[weightClass] || [];
                lotteryResults[weightClass].push({
                    athlete_id: ath.id,
                    slot: slot++,
                    day: 1,
                    name: ath.name,
                    weight: ath.weight,
                    gender: ath.gender,
                    birth_year: ath.birth_year,
                    region: ath.region,
                    coach: ath.coach
                });
            });
        });

        await API.saveLottery(tournamentId, lotteryResults);
        await this.performLottery(tournamentId);
    },

    async editBracketAthlete(tournamentId, button, weightClass, slot, currentAthleteId) {
        // Если меню уже открыто — закрываем его и выходим
        const openMenu = document.querySelector('.athlete-picker');
        if (openMenu) { openMenu.remove(); return; }

        const athletesAll = await API.getAthletes();

        // Показываем только спортсменов того же пола и ТОЧНОГО веса категории
        const wcMatch = String(weightClass).match(/^([а-яА-ЯёЁ]+)\s+([\d.]+)\s*кг/);
        const wcGender = wcMatch ? wcMatch[1] : null;
        const wcWeight = wcMatch ? parseFloat(wcMatch[2]) : null;
        let athletes = athletesAll.filter(a => {
            if (wcGender && a.gender !== wcGender) return false;
            if (wcWeight != null && a.weight != null) {
                if (Math.abs(Number(a.weight) - wcWeight) > 0.001) return false;
            }
            return true;
        });
        // Защита: если категория не распозналась — показываем всех, чтобы меню не было пустым
        if (athletes.length === 0 && wcMatch) {
            athletes = athletesAll;
        }

        // Позиционируем меню рядом с кнопкой
        const rect = button.getBoundingClientRect();
        const menu = document.createElement('div');
        menu.className = 'method-menu athlete-picker';
        menu.style.left = Math.max(8, Math.min(rect.left, window.innerWidth - 300)) + 'px';
        menu.style.top = (rect.bottom + 4) + 'px';

        const safeWc = this.escapeHtml(weightClass);
        let optionsHtml;
        if (athletes.length === 0) {
            optionsHtml = '<div class="picker-empty">Нет доступных спортсменов</div>';
        } else {
            optionsHtml = athletes.map(a =>
                `<button class="method-option" data-id="${a.id}" data-wc="${safeWc}" data-slot="${slot}">` +
                `${this.escapeHtml(a.name)} (${a.weight} кг)</button>`
            ).join('');
        }

        menu.innerHTML = `<div class="method-menu-title">Выберите спортсмена:</div>` + optionsHtml;
        document.body.appendChild(menu);

        // Подсвечиваем текущего спортсмена
        menu.querySelectorAll('.method-option').forEach(o => {
            if (parseInt(o.dataset.id) === currentAthleteId) o.classList.add('selected');
        });

        const close = () => { if (menu.parentNode) menu.remove(); };
        const onDocClick = (e) => { if (!menu.contains(e.target)) close(); };
        const onEsc = (e) => { if (e.key === 'Escape') close(); };
        setTimeout(() => {
            document.addEventListener('click', onDocClick);
            document.addEventListener('keydown', onEsc);
        }, 0);

        menu.addEventListener('click', async (ev) => {
            const opt = ev.target.closest('.method-option');
            if (!opt) return;
            close();
            document.removeEventListener('click', onDocClick);
            document.removeEventListener('keydown', onEsc);

            const newId = parseInt(opt.dataset.id);
            if (!newId || newId === currentAthleteId) return;

            try {
                // Загружаем текущую жеребьёвку и меняем спортсмена в этом слоте
                let lottery = await API.getLottery(tournamentId);
                const arr = lottery[weightClass] || [];
                const idx = arr.findIndex(x => x.slot === slot);
                if (idx === -1) {
                    toast('Слот не найден', 'error');
                    return;
                }

                // Если выбранный уже стоит в другом слоте этой категории — меняем их местами
                const dupIdx = arr.findIndex(x => x.athlete_id === newId && x.slot !== slot);
                if (dupIdx !== -1) {
                    arr[dupIdx].athlete_id = currentAthleteId;
                }
                arr[idx].athlete_id = newId;

                // Сохраняем ВСЮ жеребьёвку целиком (иначе saveLottery удалит остальные категории)
                const payload = {};
                for (const [wc, list] of Object.entries(lottery)) {
                    payload[wc] = list.map(x => ({ athlete_id: x.athlete_id, slot: x.slot }));
                }
                await API.saveLottery(tournamentId, payload);

                // Очищаем результаты этой категории, чтобы они пересчитались заново
                await API.clearWeightClassResults(tournamentId, weightClass);

                toast('✅ Соперник заменён', 'success');

                const updatedTournament = await API.getTournament(tournamentId);
                const keepDay = document.querySelector('#bracket-container .day-tab.active')?.dataset.day;
                await this.renderBracket(updatedTournament);
                if (keepDay) {
                    const cont = document.getElementById('bracket-container');
                    cont.querySelectorAll('.day-tab').forEach(t =>
                        t.classList.toggle('active', String(t.dataset.day) === String(keepDay)));
                    cont.querySelectorAll('.day-panel').forEach(p =>
                        p.classList.toggle('active', String(p.dataset.day) === String(keepDay)));
                }
                await this.renderTournaments();
                const rs = document.getElementById('results-tournament-select');
                if (rs && String(rs.value) === String(tournamentId)) {
                    await this.renderResults(updatedTournament);
                }
            } catch (err) {
                toast('Ошибка при замене: ' + err.message, 'error');
            }
        });
    },

    // ==================== ЖЕРЕБЬЁВКА: РЕДАКТИРОВАНИЕ СОПЕРНИКОВ И ДНЕЙ ====================

    // Замена спортсмена в жеребьёвке (тот же алгоритм, что и в сетке: выбор по весовой категории)
    async editLotteryAthlete(tournamentId, button, weightClass, slot, currentAthleteId) {
        const openMenu = document.querySelector('.athlete-picker');
        if (openMenu) { openMenu.remove(); return; }

        const row = button.closest('tr');
        const athletesAll = await API.getAthletes();

        // Показываем только спортсменов того же пола и ТОЧНОГО веса категории
        const wcMatch = String(weightClass).match(/^([а-яА-ЯёЁ]+)\s+([\d.]+)\s*кг/);
        const wcGender = wcMatch ? wcMatch[1] : null;
        const wcWeight = wcMatch ? parseFloat(wcMatch[2]) : null;
        let athletes = athletesAll.filter(a => {
            if (wcGender && a.gender !== wcGender) return false;
            if (wcWeight != null && a.weight != null) {
                if (Math.abs(Number(a.weight) - wcWeight) > 0.001) return false;
            }
            return true;
        });
        // Защита: если категория не распозналась — показываем всех, чтобы меню не было пустым
        if (athletes.length === 0 && wcMatch) athletes = athletesAll;

        const rect = button.getBoundingClientRect();
        const menu = document.createElement('div');
        menu.className = 'method-menu athlete-picker';
        menu.style.left = Math.max(8, Math.min(rect.left, window.innerWidth - 300)) + 'px';
        menu.style.top = (rect.bottom + 4) + 'px';

        const safeWc = this.escapeHtml(weightClass);
        const optionsHtml = athletes.length
            ? athletes.map(a =>
                `<button class="method-option" data-id="${a.id}" data-wc="${safeWc}" data-slot="${slot}">` +
                `${this.escapeHtml(a.name)} (${a.weight} кг)</button>`).join('')
            : '<div class="picker-empty">Нет доступных спортсменов</div>';
        menu.innerHTML = `<div class="method-menu-title">Выберите спортсмена:</div>` + optionsHtml;
        document.body.appendChild(menu);

        menu.querySelectorAll('.method-option').forEach(o => {
            if (parseInt(o.dataset.id) === currentAthleteId) o.classList.add('selected');
        });

        const close = () => { if (menu.parentNode) menu.remove(); };
        const onDocClick = (e) => { if (!menu.contains(e.target)) close(); };
        const onEsc = (e) => { if (e.key === 'Escape') close(); };
        setTimeout(() => {
            document.addEventListener('click', onDocClick);
            document.addEventListener('keydown', onEsc);
        }, 0);

        menu.addEventListener('click', (ev) => {
            const opt = ev.target.closest('.method-option');
            if (!opt) return;
            close();
            document.removeEventListener('click', onDocClick);
            document.removeEventListener('keydown', onEsc);

            const newId = parseInt(opt.dataset.id);
            if (!newId || newId === currentAthleteId) return;

            const table = row.closest('.lottery-table');
            let duplicateRow = null;
            if (table) {
                for (const tr of table.querySelectorAll('tbody tr')) {
                    if (tr !== row && parseInt(tr.dataset.athleteId) === newId) {
                        duplicateRow = tr; break;
                    }
                }
            }
            // Если выбранный уже стоит в другом слоте — меняем их местами (без дублей)
            if (duplicateRow) {
                const prev = athletes.find(a => a.id === currentAthleteId);
                const nc2 = duplicateRow.querySelector('td:nth-child(3)');
                const wc2 = duplicateRow.querySelector('td:nth-child(4)');
                if (nc2) { nc2.textContent = prev ? prev.name : ''; nc2.style.color = '#ffd700'; nc2.style.fontWeight = 'bold'; }
                if (wc2) { wc2.textContent = prev ? prev.weight + ' кг' : ''; wc2.style.color = '#ffd700'; }
                duplicateRow.dataset.athleteId = currentAthleteId;
            }

            const newAthlete = athletes.find(a => a.id === newId);
            const nc = row.querySelector('td:nth-child(3)');
            const wc3 = row.querySelector('td:nth-child(4)');
            if (nc) { nc.textContent = newAthlete ? newAthlete.name : ''; nc.style.color = '#ffd700'; nc.style.fontWeight = 'bold'; }
            if (wc3) { wc3.textContent = (newAthlete ? newAthlete.weight : '') + ' кг'; wc3.style.color = '#ffd700'; }
            row.dataset.athleteId = newId;
        });
    },

    // Сохранение жеребьёвки с учётом выбранных дней и заменённых соперников
    async saveLotteryFromUI() {
        const results = document.getElementById('lottery-results');
        const rows = results.querySelectorAll('tbody tr');
        const tournamentId = document.getElementById('lottery-tournament-select').value;
        if (!tournamentId) { toast('❌ Выберите соревнование!'); return; }

        const lotteryData = {};
        let hasChanges = false;
        rows.forEach(row => {
            const wc = row.dataset.weightClass;
            const athleteId = parseInt(row.dataset.athleteId);
            const slot = parseInt(row.querySelector('td strong').textContent);
            const daySelect = row.querySelector('.lottery-day-select');
            const day = daySelect ? parseInt(daySelect.value) || 1 : 1;
            if (athleteId) {
                if (!lotteryData[wc]) lotteryData[wc] = [];
                lotteryData[wc].push({ athlete_id: athleteId, slot: slot, day: day });
                hasChanges = true;
            }
        });

        if (!hasChanges) { toast('⚠️ Нет спортсменов для сохранения!'); return; }

        for (const wc of Object.keys(lotteryData)) {
            lotteryData[wc].sort((a, b) => a.slot - b.slot);
        }

        try {
            await API.saveLottery(parseInt(tournamentId), lotteryData);
            toast('✅ Жеребьевка сохранена!');

            // Перерисовываем сетку и результаты, если показан тот же турнир
            const bs = document.getElementById('bracket-tournament-select');
            const rs = document.getElementById('results-tournament-select');
            const tId = parseInt(tournamentId);
            if (bs && String(bs.value) === String(tId)) {
                const updatedTournament = await API.getTournament(tId);
                await this.renderBracket(updatedTournament);
            }
            if (rs && String(rs.value) === String(tId)) {
                const updatedTournament = await API.getTournament(tId);
                await this.renderResults(updatedTournament);
            }
        } catch (err) {
            toast('❌ Ошибка сохранения: ' + err.message);
        }
    },

    // ==================== СЕТКА ====================

    async renderBracketSelect() {
        const select = document.getElementById('bracket-tournament-select');
        const container = document.getElementById('bracket-container');
        const classesSelector = document.getElementById('weight-classes-selector');
        const classesList = document.getElementById('weight-classes-list');

        if (!select) return;

        // Загружаем список соревнований
        try {
            const tournaments = await API.getTournaments();
            select.innerHTML = '<option value="">-- Выберите соревнование --</option>';
            tournaments.forEach(t => {
                const opt = document.createElement('option');
                opt.value = t.id;
                opt.textContent = `${t.name} (${this.fmtDate(t.date_from || t.date_to) || '—'})`;
                select.appendChild(opt);
            });

            // Загружаем сохраненное соревнование по умолчанию
            let defaultId = this.getDefaultTournamentId();
            
            // Если defaultId не установлен, но есть турниры - выбираем первый
            if (!defaultId && tournaments.length > 0) {
                defaultId = tournaments[0].id.toString();
            }
            
            if (defaultId) {
                const opt = select.querySelector(`option[value="${defaultId}"]`);
                if (opt) opt.selected = true;
                
                // Загружаем данные турнира
                const tournament = await API.getTournament(parseInt(defaultId));
                if (tournament.athlete_ids && tournament.athlete_ids.length > 0) {
                    await this.loadWeightClasses(parseInt(defaultId));
                    // Не вызываем renderBracket здесь — он вызовется в loadBracketFromDB
                } else {
                    container.innerHTML = '<div class="empty-state"><p>Выберите соревнование для просмотра сетки</p></div>';
                }
            }

            // Привязываем обработчики только один раз
            if (!select.dataset.initialized) {
                select.dataset.initialized = 'true';

                select.addEventListener('change', async () => {
                    const id = parseInt(select.value);
                    if (id) {
                        this.saveDefaultTournament(id);
                        const tournament = await API.getTournament(id);
                        
                        if (!tournament.athlete_ids || tournament.athlete_ids.length === 0) {
                            classesSelector.style.display = 'none';
                            container.innerHTML = `
                                <div class="empty-state">
                                    <p>⚠️ В соревновании нет участников!</p>
                                    <p>Добавьте спортсменов через кнопку <strong>"👥 Участники"</strong> в разделе "Соревнования"</p>
                                </div>
                            `;
                            return;
                        }
                        
                        await this.loadWeightClasses(id);
                        await this.renderBracket(tournament);
                    } else {
                        this.saveDefaultTournament(null);
                        classesSelector.style.display = 'none';
                        container.innerHTML = '<div class="empty-state"><p>Выберите соревнование для просмотра сетки</p></div>';
                    }
                });

                // Кнопка "Сформировать сетку"
                const generateBtn = document.getElementById('generate-bracket');
                if (generateBtn) {
                    generateBtn.addEventListener('click', async () => {
                        const id = parseInt(select.value);
                        if (!id) {
                            toast('Выберите соревнование!');
                            return;
                        }
                        try {
                            // Просто загружаем турнир и рендерим сетку из lottery
                            const tournament = await API.getTournament(id);
                            await this.renderBracket(tournament);
                        } catch (err) {
                            toast('Ошибка: ' + err.message);
                        }
                    });
                }

                const selectAllBtn = document.getElementById('select-all-weight-classes');
                if (selectAllBtn) {
                    selectAllBtn.addEventListener('click', () => {
                        classesList.querySelectorAll('input').forEach(c => c.checked = true);
                    });
                }
            }
        } catch (err) {
            console.error('Ошибка:', err);
        }
    },

    async loadWeightClasses(tournamentId) {
        const selector = document.getElementById('weight-classes-selector');
        const list = document.getElementById('weight-classes-list');
        // Блок выбора категорий отключён (закомментирован в HTML) — выходим
        if (!selector || !list) return;
        try {
            const { weight_classes } = await API.getWeightClasses(tournamentId);
            if (!weight_classes || weight_classes.length === 0) {
                selector.style.display = 'none';
                return;
            }
            list.innerHTML = weight_classes.map((wc, i) => `
                <label class="weight-class-check">
                    <input type="checkbox" value="${this.escapeHtml(wc)}" checked>
                    ${this.escapeHtml(wc)} (${i === 0 ? 'по умолчанию' : ''})
                </label>
            `).join('');
            selector.style.display = 'block';
        } catch (err) {
            console.error('Ошибка загрузки категорий:', err);
            selector.style.display = 'none';
        }
    },

    async renderBracket(tournament) {
        const container = document.getElementById('bracket-container');
        if (!container) return;

        // Возвращает ID бойца из объекта слота или числового id
        const fid = (x) => (x && typeof x === 'object') ? x.athlete_id : x;

        // Запоминаем ID турнира для обработчика кнопок победителя
        this.currentBracketTournamentId = tournament.id;

        try {
            // Получаем жеребьевку из БД (объект: { "категория": [бойцы], ... })
            const [lottery, results] = await Promise.all([
                API.getLottery(tournament.id),
                API.getBracketResults(tournament.id).catch(() => [])
            ]);

            // Кэш спортсменов — для имён победителей в следующих раундах
            this.athletesCache = await API.getAthletes();

            if (!lottery || typeof lottery !== 'object' || Object.keys(lottery).length === 0) {
                container.innerHTML = '<div class="empty-state"><p>Жеребьевка не проведена. Проведите жеребьевку во вкладке "Жеребьевка"</p></div>';
                return;
            }

            // Карта сохранённых результатов: "категория|индекс_пары" -> {winner_id, method}
            const resultsMap = {};
            (Array.isArray(results) ? results : []).forEach(r => {
                resultsMap[`${r.weight_class}|${r.pair_index}`] = r;
            });

            // Способы победы
            const WIN_METHODS = ['Нокаут', 'Технический нокаут', 'Единогласно', 'Раздельное решение', 'По очкам', 'Дисквалификация', 'Отказ от боя', 'Травма'];

            // Сортируем весовые категории: сначала по полу, потом по весу
            function getWeightLower(wc) {
                const match = wc.match(/(\d+(?:\.\d+)?)/);
                return match ? parseFloat(match[1]) : 999;
            }
            
            const sortedWeightClasses = Object.keys(lottery).sort((a, b) => {
                const genderA = a.split(' ')[0];
                const genderB = b.split(' ')[0];
                const weightA = getWeightLower(a);
                const weightB = getWeightLower(b);
                if (genderA !== genderB) return genderA.localeCompare(genderB);
                return weightA - weightB;
            });

// === Сборка сетки с разбивкой по дням турнира ===
            const daysMap = {};        // день -> { cat: { весовая категория -> { slotCount, cols } } }
            const championByWc = {};   // весовая категория -> чемпион (null/объект)
            const silverByWc = {};     // весовая категория -> серебряный призёр (имя/null)
            // Количество дней: берём максимум из поля days и интервала дат (date_from..date_to)
            let daysFromDates = 1;
            if (tournament.date_from && tournament.date_to) {
                const diff = (new Date(tournament.date_to) - new Date(tournament.date_from)) / 86400000;
                if (isFinite(diff) && diff >= 0) daysFromDates = Math.floor(diff) + 1;
            }
            const tournamentDays = Math.max(daysFromDates, Math.max(1, parseInt(tournament.days) || 1));
            for (const weightClass of sortedWeightClasses) {
                // Сортируем бойцов по слоту (slot из жеребьевки)
                const slots = [...lottery[weightClass]].sort((a, b) => a.slot - b.slot);

                // Строим раунды: раунд 1 — пары из жеребьевки, далее — победители
                // rounds = [ [ {a1, a2, pairIndex, result}, ... ], [ ... ], ... ]
                const rounds = [];
                let firstRound = [];
                let pairIndex = 0;
                for (let i = 0; i < slots.length - 1; i += 2) {
                    firstRound.push({
                        a1: slots[i],
                        a2: slots[i + 1],
                        pairIndex: pairIndex,
                        result: resultsMap[`${weightClass}|${pairIndex}`] || null
                    });
                    pairIndex++;
                }
                if (slots.length % 2 !== 0) {
                    // Нечётный боец — автоматический проход
                    firstRound.push({
                        a1: slots[slots.length - 1],
                        a2: null,
                        pairIndex: null,
                        result: { winner_id: slots[slots.length - 1].athlete_id, method: 'авто' }
                    });
                }
                rounds.push(firstRound);

                // Следующие раунды: пары победителей.
                // pairIndex для раундов 2+ — раунд-кодированный (1000*round + idx),
                // чтобы результат каждой пары можно было сохранить в БД.
                let roundNum = 1;
                while (true) {
                    const prev = rounds[rounds.length - 1];
                    const winners = prev.map(f => f.result ? f.result.winner_id : null);
                    // Если в предыдущем раунде ещё не все бои завершены — дальше не строим
                    if (winners.some(w => w === null)) break;
                    if (winners.length <= 1) break;
                    roundNum++;
                    const offset = (roundNum - 1) * 1000;
                    const nextRound = [];
                    for (let i = 0; i < winners.length - 1; i += 2) {
                        const pIdx = offset + i;
                        nextRound.push({
                            a1: winners[i],
                            a2: winners[i + 1],
                            pairIndex: pIdx,
                            result: resultsMap[`${weightClass}|${pIdx}`] || null
                        });
                    }
                    if (nextRound.length === 0) break;
                    rounds.push(nextRound);
                }

                const maxRounds = rounds.length; 
                // tournamentDays объявлен выше (с учётом интервала дат) — здесь не переобъявляем
                // const tournamentDays = Math.max(1, parseInt(tournament.days) || 1);

                // Авто-публикация результата категории:
                // чемпион определён, если не осталось нерешённых пар (нет соперников).
                let championId = null;
                const anyUnresolved = rounds.some(r => r.some(f => f.a2 !== null && !f.result));
                const top = rounds[rounds.length - 1];
                const topWinners = top.map(f => f.result ? f.result.winner_id : null);
                if (!anyUnresolved && topWinners.length === 1 && topWinners[0] != null) {
                    championId = topWinners[0];
                }
                const champion = championId != null ? {
                    athlete_id: championId,
                    name: this.getAthleteName(championId)
                } : null;

                // Запоминаем чемпиона категории (нужен в блоках дней)
                championByWc[weightClass] = champion;

                // Серебряный призёр — проигравший в финале (если финал сыгран)
                let silverName = null;
                {
                    const fr = rounds[rounds.length - 1];
                    if (champion && fr && fr.length === 1 && fr[0].a2 !== null && fr[0].result) {
                        const fp = fr[0];
                        const la = fid(fp.a1), lb = fid(fp.a2);
                        const loserId = (fp.result.winner_id === la) ? lb : la;
                        if (loserId != null) silverName = this.getAthleteName(loserId);
                    }
                }
                silverByWc[weightClass] = silverName || null;

                rounds.forEach((roundFights, roundIdx) => {
                    const roundNum = roundIdx + 1;
                    const roundName = this.getRoundName(roundNum, maxRounds);
                    // Каждый БОЙ получает свой день:
                    //  - 1-й раунд — из жеребьёвки (день слота), ограниченный числом дней
                    //  - последующие раунды — по номеру раунда (когда сыграны предыдущие)
                    roundFights.forEach(f => {
                        let dayNum;
                        if (roundNum === 1 && f.a1 && typeof f.a1 === 'object' && f.a1.day != null) {
                            dayNum = Math.min(tournamentDays, Math.max(1, parseInt(f.a1.day) || 1));
                        } else {
                            dayNum = this.roundToDay(roundNum, maxRounds, tournamentDays);
                        }
                        if (!daysMap[dayNum]) daysMap[dayNum] = { cat: {} };
                        if (!daysMap[dayNum].cat[weightClass]) {
                            daysMap[dayNum].cat[weightClass] = { slotCount: slots.length, cols: {} };
                        }
                        // Колонка раунда создаётся на день — ключ по номеру раунда
                        const colKey = String(roundNum);
                        if (!daysMap[dayNum].cat[weightClass].cols[colKey]) {
                            daysMap[dayNum].cat[weightClass].cols[colKey] = `<div class="bracket-round-col" data-day="${dayNum}">
                                <div class="round-label"><span class="day-tag">День ${dayNum}</span> ${roundName}</div>`;
                        }

                        // Разрешаем и объекты бойцов, и ID (из победителей предыдущего раунда)
                        const a1 = (typeof f.a1 === 'object') ? f.a1 : { athlete_id: f.a1, name: this.getAthleteName(f.a1), weight: '' };
                        const a2 = f.a2 === null ? null : (typeof f.a2 === 'object') ? f.a2 : { athlete_id: f.a2, name: this.getAthleteName(f.a2), weight: '' };

                        const savedWinner = f.result ? f.result.winner_id : null;
                        const savedMethod = f.result ? f.result.method : '';

                        const a1Class = savedWinner != null && savedWinner === a1.athlete_id ? 'winner' :
                                        savedWinner != null ? 'loser' : '';
                        const a2Class = a2 === null ? 'empty' :
                                        savedWinner != null && savedWinner === a2.athlete_id ? 'winner' :
                                        savedWinner != null ? 'loser' : '';

                        // Кнопки выбора победителя — только если есть соперник и результат не выбран
                        let buttonsHtml = '';
                        if (a2 !== null && savedWinner == null && f.pairIndex !== null) {
                            buttonsHtml = `
                                <div class="fighter-buttons">
                                    <button class="btn-winner btn-winner-1" data-winner="${a1.athlete_id}">
                                        ✓ Победитель: ${this.escapeHtml(a1.name)}
                                    </button>
                                    <button class="btn-winner btn-winner-2" data-winner="${a2.athlete_id}">
                                        ✓ Победитель: ${this.escapeHtml(a2.name)}
                                    </button>
                                </div>`;
                        }

                        const methodLine = savedWinner != null && savedMethod && savedMethod !== 'авто'
                            ? `<div class="fight-method">${this.escapeHtml(savedMethod)}</div>` : '';
                        const byeLine = a2 === null ? '<div class="fight-method">Автоматический проход</div>' : '';

                        daysMap[dayNum].cat[weightClass].cols[colKey] += `
                            <div class="bracket-fight" data-pair-index="${f.pairIndex !== null ? f.pairIndex : ''}" data-day="${dayNum}">
                                <div class="fighter-row ${a1Class}" data-slot="${a1.slot != null ? a1.slot : ''}">
                                    <span class="fighter-name">${this.escapeHtml(a1.name)}${a1.weight ? ` (${a1.weight} кг)` : ''}</span>
                                    ${savedWinner === a1.athlete_id ? '<span class="crown">👑</span>' : ''}
                                    ${roundNum === 1 && a1.slot != null ? `<button class="btn-edit-fighter" onclick="App.editBracketAthlete(${tournament.id}, this, '${this.escapeHtml(weightClass)}', ${a1.slot}, ${a1.athlete_id})" title="Заменить спортсмена">✏️</button>` : ''}
                                </div>
                                <div class="fighter-row ${a2Class}" data-slot="${a2 !== null && a2.slot != null ? a2.slot : ''}">
                                    <span class="fighter-name">${a2 === null ? 'BYE' : this.escapeHtml(a2.name) + (a2.weight ? ` (${a2.weight} кг)` : '')}</span>
                                    ${a2 !== null && savedWinner === a2.athlete_id ? '<span class="crown">👑</span>' : ''}
                                    ${roundNum === 1 && a2 !== null && a2.slot != null ? `<button class="btn-edit-fighter" onclick="App.editBracketAthlete(${tournament.id}, this, '${this.escapeHtml(weightClass)}', ${a2.slot}, ${a2.athlete_id})" title="Заменить спортсмена">✏️</button>` : ''}
                                </div>
                                ${buttonsHtml}
                                ${methodLine}
                                ${byeLine}
                            </div>`;
                        // Ячейка победителя в финале (бой за 1-2 места) — внизу слота, размером со слот
                        if (roundNum === maxRounds && savedWinner != null && a2 !== null) {
                            const wname = (savedWinner === a1.athlete_id) ? a1.name : a2.name;
                            daysMap[dayNum].cat[weightClass].cols[colKey] +=
                                `<div class="fight-result-cell">Победитель: ${this.escapeHtml(wname)}</div>`;
                        }
                    });
                });

                // Бронзовый бой за 3-е/4-е место ОТКЛЮЧЁН — бронзу получают оба полуфиналиста.
                // (Блок оставлен закомментированным на случай, если решите вернуть его обратно.)
                /* if (maxRounds >= 2 && rounds.length >= 2) {
                    const semi = rounds[rounds.length - 2];
                    const finals = rounds[rounds.length - 1];
                    if (semi.length === 2 && semi.every(f => f.result && f.a2 !== null) && finals.length === 1) {
                        const losers = semi.map(f => {
                            const la = fid(f.a1), lb = fid(f.a2);
                            return (f.result.winner_id === la) ? lb : la;
                        });
                        if (losers[0] != null && losers[1] != null && losers[0] !== losers[1]) {
                            const bPair = 9000;
                            const bRes = resultsMap[`${weightClass}|${bPair}`] || null;
                            const bDay = this.roundToDay(maxRounds, maxRounds, tournamentDays);
                            if (!daysMap[bDay]) daysMap[bDay] = { cat: {} };
                            if (!daysMap[bDay].cat[weightClass]) daysMap[bDay].cat[weightClass] = { slotCount: slots.length, cols: {} };
                            const bKey = String(maxRounds + 1);
                            daysMap[bDay].cat[weightClass].cols[bKey] = `<div class="bracket-round-col" data-day="${bDay}">
                                <div class="round-label"><span class="day-tag">День ${bDay}</span> Бронзовый бой</div>`;

                            const fb = losers[0], fs = losers[1];
                            const bn1 = this.getAthleteName(fb), bn2 = this.getAthleteName(fs);
                            const bw = bRes ? bRes.winner_id : null;
                            const bm = bRes ? bRes.method : '';
                            const b1c = bw != null && bw === fb ? 'winner' : bw != null ? 'loser' : '';
                            const b2c = bw != null && bw === fs ? 'winner' : bw != null ? 'loser' : '';
                            let bbuttons = '';
                            if (bw == null) {
                                bbuttons = `
                                    <div class="fighter-buttons">
                                        <button class="btn-winner btn-winner-1" data-winner="${fb}">✓ Победитель: ${this.escapeHtml(bn1)}</button>
                                        <button class="btn-winner btn-winner-2" data-winner="${fs}">✓ Победитель: ${this.escapeHtml(bn2)}</button>
                                    </div>`;
                            }
                            const bmethod = bw != null && bm ? `<div class="fight-method">${this.escapeHtml(bm)}</div>` : '';
                            daysMap[bDay].cat[weightClass].cols[bKey] += `
                                <div class="bracket-fight" data-pair-index="${bPair}" data-day="${bDay}">
                                    <div class="fighter-row ${b1c}">
                                        <span class="fighter-name">${this.escapeHtml(bn1)}</span>
                                        ${bw === fb ? '<span class="crown">👑</span>' : ''}
                                    </div>
                                    <div class="fighter-row ${b2c}">
                                        <span class="fighter-name">${this.escapeHtml(bn2)}</span>
                                        ${bw === fs ? '<span class="crown">👑</span>' : ''}
                                    </div>
                                    ${bbuttons}
                                    ${bmethod}
                                </div>`;
                            if (bw != null) {
                                const bwin = (bw === fb) ? bn1 : bn2;
                                daysMap[bDay].cat[weightClass].cols[bKey] +=
                                    `<div class="fight-result-cell">Победитель: ${this.escapeHtml(bwin)}</div>`;
                            }
                        }
                    }
                } */
            }

            // Закрываем колонки раундов и приводим к отсортированному по раунду массиву
            for (const dayKey of Object.keys(daysMap)) {
                const dag = daysMap[dayKey];
                for (const wc of Object.keys(dag.cat)) {
                    const c = dag.cat[wc];
                    const orderedCols = [];
                    Object.keys(c.cols).sort((a, b) => a - b).forEach(k => orderedCols.push(c.cols[k] + '</div>'));
                    c.cols = orderedCols;
                }
            }

            // === Отрисовка вкладок дней и их содержимого ===
            // Показываем все дни интервала (даже пустые), чтобы была видна сетка расписания
            const orderedDays = [];
            for (let d = 1; d <= tournamentDays; d++) orderedDays.push(d);
            const emptyDay = { cat: {} };

            // Последний день каждой категории (там финал) — баннер показываем только там
            const maxDayByWc = {};
            for (const dKey of Object.keys(daysMap)) {
                const d = parseInt(dKey);
                for (const wc of Object.keys(daysMap[dKey].cat)) {
                    maxDayByWc[wc] = Math.max(maxDayByWc[wc] || 0, d);
                }
            }

            let html = '<div class="bracket-days">';
            html += '<div class="day-tabs">' + orderedDays.map((d, i) =>
                `<button class="day-tab${i === 0 ? ' active' : ''}" data-day="${d}">День ${d}</button>`
            ).join('') + '</div>';

            let panels = '';
            orderedDays.forEach((d, i) => {
                const day = daysMap[d] || emptyDay;
                let body = '';
                // Категории этого дня — отсортированы так же, как в общем списке
                Object.keys(day.cat).sort((a, b) => {
                    const gA = a.split(' ')[0], gB = b.split(' ')[0];
                    const wA = getWeightLower(a), wB = getWeightLower(b);
                    if (gA !== gB) return gA.localeCompare(gB);
                    return wA - wB;
                }).forEach(wc => {
                    const c = day.cat[wc];
                    let catBody = `<div class="bracket-weight-class">
                        <h3 class="weight-title" data-weight-class="${this.escapeHtml(wc)}">${this.escapeHtml(wc)} (${c.slotCount} бойцов)</h3>
                        <div class="bracket-tree">${c.cols.join('')}</div>`;

                    const champ = championByWc[wc];
                    // Баннер победителя в сетке убран — победители видны в ячейках под слотами
                    // (смотрим fight-result-cell под финалом и бронзовым боем).
                    // Блок оставлен закомментированным на всякий случай.
                    /* if (champ && maxDayByWc[wc] === d) {
                        catBody += `
                            <div class="weight-championship">
                                <span class="wt-champ-crown">👑</span>
                                <div>
                                    <div class="wt-champ-title">Победитель</div>
                                    <div class="wt-champ-name">${this.escapeHtml(champ.name)}</div>
                                </div>
                            </div>`;
                    } */
                    catBody += '</div>';
                    body += catBody;
                });
                panels += `<div class="day-panel${i === 0 ? ' active' : ''}" data-day="${d}">${body}</div>`;
            });

            html += panels + '</div>';
            container.innerHTML = html;

            // Обработчик переключения вкладок дней
            if (!container.dataset.dayHandler) {
                container.dataset.dayHandler = 'true';
                container.addEventListener('click', (e) => {
                    const tab = e.target.closest('.day-tab');
                    if (!tab) return;
                    const d = tab.dataset.day;
                    container.querySelectorAll('.day-tab').forEach(t => t.classList.toggle('active', t === tab));
                    container.querySelectorAll('.day-panel').forEach(p => p.classList.toggle('active', p.dataset.day === d));
                });
            }

            // Обработчик клика по кнопкам победителя (делегирование)
            if (!container.dataset.winnerHandler) {
                container.dataset.winnerHandler = 'true';
                container.addEventListener('click', async (e) => {
                    const btn = e.target.closest('.btn-winner');
                    if (!btn) return;
                    const fightEl = btn.closest('.bracket-fight');
                    if (!fightEl) return;

                    // Если меню уже открыто в этом бою:
                    //  - клик по той же кнопке — закрываем меню
                    //  - клик по другой кнопке — переоткрываем для нового победителя
                    const existingMenu = fightEl.querySelector('.method-menu');
                    if (existingMenu) {
                        const sameButton = existingMenu.dataset.winnerId === btn.dataset.winner;
                        existingMenu.remove();
                        if (sameButton) return;
                    }

                    const winnerId = parseInt(btn.dataset.winner);
                    const pairIndex = parseInt(fightEl.dataset.pairIndex);
                    const weightBlock = fightEl.closest('.bracket-weight-class');
                    const weightClass = weightBlock.querySelector('.weight-title').dataset.weightClass;

                    // Выпадающее меню выбора способа победы
                    const menu = document.createElement('div');
                    menu.className = 'method-menu';
                    menu.dataset.winnerId = btn.dataset.winner;
                    menu.innerHTML = `<div class="method-menu-title">Способ победы:</div>` +
                        WIN_METHODS.map(m =>
                            `<button class="method-option" data-method="${this.escapeHtml(m)}">${this.escapeHtml(m)}</button>`
                        ).join('');
                    fightEl.appendChild(menu);

                    menu.addEventListener('click', async (ev) => {
                        const opt = ev.target.closest('.method-option');
                        if (!opt) return;
                        const method = opt.dataset.method;
                        menu.remove();
                        try {
                            const day = parseInt(fightEl.dataset.day) || 1;
                            await API.saveBracketResult(this.currentBracketTournamentId, weightClass, pairIndex, winnerId, method, day);

                            // Публикация результата: если категория стала полностью решённой
                            // (чемпион определён) — показываем уведомление и обновляем «Результаты».
                            const championBefore = new Set(this._publishedChampions || []);
                            // Сохраняем активный день, чтобы не сбрасывался на первый после перерисовки
                            const keepDay = container.querySelector('.day-tab.active')?.dataset.day;
                            const tournament = await API.getTournament(this.currentBracketTournamentId);
                            await this.renderBracket(tournament);
                            if (keepDay) {
                                container.querySelectorAll('.day-tab').forEach(t =>
                                    t.classList.toggle('active', String(t.dataset.day) === String(keepDay)));
                                container.querySelectorAll('.day-panel').forEach(p =>
                                    p.classList.toggle('active', String(p.dataset.day) === String(keepDay)));
                            }
                            const championsNow = [...document.querySelectorAll('.weight-championship .wt-champ-name')]
                                .map(n => n.textContent.trim());
                            const newChampions = championsNow.filter(c => !championBefore.has(c));
                            this._publishedChampions = championsNow;
                            if (newChampions.length > 0) {
                                toast('🏆 Категория завершена! Победитель опубликован в результатах: ' + newChampions[0], 'success', 5000);
                                const resultsTab = document.getElementById('tab-results');
                                if (resultsTab) {
                                    const rs = document.getElementById('results-tournament-select');
                                    if (rs && String(rs.value) === String(this.currentBracketTournamentId)) {
                                        await this.renderResults(tournament);
                                    }
                                }
                            }
                        } catch (err) {
                            toast('Ошибка сохранения победителя: ' + err.message);
                        }
                    });
                });

                // Клик вне меню (и вне кнопок победителя) закрывает все открытые меню
                document.addEventListener('click', (e) => {
                    if (e.target.closest('.method-menu') || e.target.closest('.btn-winner')) return;
                    document.querySelectorAll('.method-menu').forEach(m => m.remove());
                });

                // Esc закрывает все открытые меню
                document.addEventListener('keydown', (e) => {
                    if (e.key === 'Escape') {
                        document.querySelectorAll('.method-menu').forEach(m => m.remove());
                    }
                });
            }
        } catch (err) {
            console.error('Ошибка рендеринга сетки:', err);
            container.innerHTML = `<div class="empty-state"><p>Ошибка: ${err.message}</p></div>`;
        }
    },

    // Имя спортсмена по ID (для последующих раундов сетки)
    getAthleteName(athleteId) {
        if (athleteId == null) return '—';
        const a = (this.athletesCache || []).find(x => x.id === athleteId);
        return a ? a.name : `#${athleteId}`;
    },

    getRoundName(round, maxRounds) {
        const names = {
            1: '1/8 финала',
            2: '1/4 финала',
            3: '1/2 финала',
            4: 'Финал',
        };
        if (names[round]) return names[round];
        // Последний раунд — всегда Финал
        if (round === maxRounds) return 'Финал';
        // Предпоследний — Победитель (если это не финал)
        if (round === maxRounds - 1) return 'Победитель';
        return `${round} раунд`;
    },

    // Номер дня для раунда: каждый раунд проходит в свой день
    // 1-й раунд — день 1, 2-й раунд — день 2 и т.д.
    // Если турнир задан на меньшее число дней — раунды сжимаются в эти дни.
    roundToDay(round, maxRounds, totalDays) {
        const days = Math.max(1, parseInt(totalDays) || 1);
        if (!maxRounds || maxRounds <= 0) return 1;
        if (days >= maxRounds) return round; // каждый раунд — отдельный день
        return Math.min(days, 1 + Math.floor((round - 1) * days / maxRounds));
    },

    // Раунд по сквозному pairIndex (раунд-кодирован: 1-й раунд = 0.., 2-й = 1000..)
    roundFromPairIndex(pairIndex) {
        if (pairIndex == null) return 1;
        return Math.floor(parseInt(pairIndex) / 1000) + 1;
    },

    // ==================== РУЧНАЯ ЗАМЕНА СПОРТСМЕНА В СЕТКЕ ====================

    async editFighter(tournamentId, fightId, slot) {
        // Проверяем, не открыт ли уже какой-то выбор
        if (document.querySelector('.fighter-select')) {
            return;
        }

        const fightEl = document.querySelector(`.bracket-fight[data-fight-id="${fightId}"]`);
        if (!fightEl) return;
        const row = fightEl.querySelector(`.fighter-row[data-slot="${slot}"]`);
        if (!row) return;

        const athletes = await API.getAthletes();
        const currentId = parseInt(row.dataset.athleteId) || 0;

        let options = '<option value="">— Выберите —</option>';
        athletes.forEach(a => {
            const selected = a.id === currentId ? 'selected' : '';
            options += `<option value="${a.id}" ${selected}>${this.escapeHtml(a.name)} (${a.weight} кг)</option>`;
        });

        row.innerHTML = `<select class="fighter-select" data-tournament="${tournamentId}" data-fight="${fightId}" data-slot="${slot}">
            ${options}
        </select>`;

        const select = row.querySelector('select');
        select.focus();
        select.addEventListener('change', async () => {
            const newId = parseInt(select.value) || null;
            try {
                await API.updateFightAthlete(tournamentId, fightId, slot, newId);
                // Если новый спортсмен встал в этот бой — перерисовываем сетку
                const updatedTournament = await API.getTournament(tournamentId);
                await this.renderBracket(updatedTournament);
            } catch (err) {
                toast('Ошибка при замене спортсмена: ' + err.message);
            }
        });

        select.addEventListener('blur', () => {
            // Если пользователь не выбрал (кликнул мимо), отменяем через перерисовку
            // Не делаем ничего — change сработает сам при выборе
        });
    },

    // ==================== РЕДАКТИРОВАНИЕ РЕЗУЛЬТАТА БОЯ ====================

    async editFightResult(tournamentId, fightId) {
        const fightEl = document.querySelector(`.bracket-fight[data-fight-id="${fightId}"]`);
        if (!fightEl) return;

        const tournament = await API.getTournament(tournamentId);
        const fight = tournament.fights.find(f => f.id === fightId);
        const athletes = await API.getAthletes();

        if (!fight) return;

        const a1 = athletes.find(a => a.id === fight.athlete1_id);
        const a2 = athletes.find(a => a.id === fight.athlete2_id);

        if (!a1 || !a2) {
            toast('❌ В бою должны быть оба спортсмена!');
            return;
        }

        // Ищем метод победы и заменяем его на выбор
        const methodDiv = fightEl.querySelector('.fight-method');
        if (!methodDiv) return;

        let options = `<option value="">— Выберите —</option>
            <option value="${a1.id}" ${fight.winner_id === a1.id ? 'selected' : ''}>${a1.name}</option>
            <option value="${a2.id}" ${fight.winner_id === a2.id ? 'selected' : ''}>${a2.name}</option>`;

        methodDiv.innerHTML = `<select class="fighter-select" data-tournament="${tournamentId}" data-fight="${fightId}" onchange="App.submitFightResult(this)">
            ${options}
        </select>`;

        const select = methodDiv.querySelector('select');
        select.focus();
    },

    async submitFightResult(selectElement) {
        const fightId = parseInt(selectElement.dataset.fight);
        const tournamentId = parseInt(selectElement.dataset.tournament);
        const winnerId = parseInt(selectElement.value);

        if (!winnerId) return;

        try {
            const tournament = await API.getTournament(tournamentId);
            const fight = tournament.fights.find(f => f.id === fightId);

            if (!fight) return;

            // Определяем метод победы
            let method = 'единогласное';

            await API.updateFight(fightId, {
                winner_id: winnerId,
                method: method,
                completed: 1
            });

            // Продвигаем победителя дальше
            await this.advanceWinner(tournamentId, fight, winnerId);

            // Перерисовываем сетку
            const updatedTournament = await API.getTournament(tournamentId);
            await this.renderBracket(updatedTournament);
            await this.renderTournaments();
        } catch (err) {
            toast('Ошибка: ' + err.message);
        }
    },

    // ==================== МОДАЛЬНОЕ ОКНО БОЯ ====================

    async openFightModal(tournamentId, fightId) {
        this.currentTournamentId = tournamentId;
        this.currentFightId = fightId;

        const tournament = await API.getTournament(tournamentId);
        const fight = tournament.fights.find(f => f.id === fightId);
        const athletes = await API.getAthletes();

        if (!fight) return;

        const a1 = athletes.find(a => a.id === fight.athlete1_id);
        const a2 = athletes.find(a => a.id === fight.athlete2_id);

        document.getElementById('fight-title').textContent =
            `Бой: ${a1 ? a1.name : 'Ожидает'} VS ${a2 ? a2.name : 'Ожидает'} (${fight.weight_class})`;

        document.getElementById('fight-details').innerHTML = `
            <p><strong>${this.escapeHtml(a1 ? a1.name : 'Ожидает')}</strong> vs <strong>${this.escapeHtml(a2 ? a2.name : 'Ожидает')}</strong></p>
            <p>Весовая категория: ${fight.weight_class}</p>
            <p>Раунд: ${this.getRoundName(fight.round, 0)}</p>
        `;

        const select = document.getElementById('fight-winner');
        select.innerHTML = '<option value="">-- Выберите победителя --</option>';
        
        if (a1) {
            const opt1 = document.createElement('option');
            opt1.value = a1.id;
            opt1.textContent = a1.name;
            select.appendChild(opt1);
        }
        if (a2) {
            const opt2 = document.createElement('option');
            opt2.value = a2.id;
            opt2.textContent = a2.name;
            select.appendChild(opt2);
        }

        // Если нет доступных спортсменов — показываем сообщение
        if (!a1 && !a2) {
            toast('⚠️ В этом бою ещё нет спортсменов!');
            document.getElementById('fight-modal').style.display = 'none';
            return;
        }

        document.getElementById('fight-modal').style.display = 'block';
    },

    setupFightResultForm() {
        // Обработчик закрытия модального окна
        const closeBtn = document.querySelector('.close');
        if (closeBtn) {
            closeBtn.addEventListener('click', () => {
                document.getElementById('fight-modal').style.display = 'none';
            });
        }

        window.addEventListener('click', (e) => {
            if (e.target === document.getElementById('fight-modal')) {
                document.getElementById('fight-modal').style.display = 'none';
            }
        });
    },

    async openFightModal(tournamentId, fightId) {
        this.currentTournamentId = tournamentId;
        this.currentFightId = fightId;

        const tournament = await API.getTournament(tournamentId);
        const fight = tournament.fights.find(f => f.id === fightId);
        const athletes = await API.getAthletes();

        if (!fight) return;

        const a1 = athletes.find(a => a.id === fight.athlete1_id);
        const a2 = athletes.find(a => a.id === fight.athlete2_id);

        document.getElementById('fight-title').textContent =
            `Бой: ${a1 ? a1.name : 'Ожидает'} VS ${a2 ? a2.name : 'Ожидает'} (${fight.weight_class})`;

        document.getElementById('fight-details').innerHTML = `
            <p><strong>${this.escapeHtml(a1 ? a1.name : 'Ожидает')}</strong> vs <strong>${this.escapeHtml(a2 ? a2.name : 'Ожидает')}</strong></p>
            <p>Весовая категория: ${fight.weight_class}</p>
            <p>Раунд: ${this.getRoundName(fight.round, 0)}</p>
        `;

        const select = document.getElementById('fight-winner');
        select.innerHTML = '<option value="">-- Выберите победителя --</option>';
        
        if (a1) {
            const opt1 = document.createElement('option');
            opt1.value = a1.id;
            opt1.textContent = a1.name;
            select.appendChild(opt1);
        }
        if (a2) {
            const opt2 = document.createElement('option');
            opt2.value = a2.id;
            opt2.textContent = a2.name;
            select.appendChild(opt2);
        }

        // Если нет доступных спортсменов — показываем сообщение
        if (!a1 && !a2) {
            toast('⚠️ В этом бою ещё нет спортсменов!');
            document.getElementById('fight-modal').style.display = 'none';
            return;
        }

        // Добавляем обработчик отправки формы
        const form = document.getElementById('fight-result-form');
        form.onsubmit = async (e) => {
            e.preventDefault();
            const winnerId = parseInt(document.getElementById('fight-winner').value);
            const method = document.getElementById('fight-method').value;

            if (!winnerId) {
                toast('Выберите победителя!');
                return;
            }

            try {
                const t = await API.getTournament(this.currentTournamentId);
                const f = t.fights.find(f => f.id === this.currentFightId);

                if (f) {
                    await API.updateFight(f.id, {
                        winner_id: winnerId,
                        method: method,
                        completed: 1
                    });

                    await this.advanceWinner(tournamentId, f, winnerId);

                    document.getElementById('fight-modal').style.display = 'none';
                    const updatedTournament = await API.getTournament(this.currentTournamentId);
                    await this.renderBracket(updatedTournament);
                    await this.renderTournaments();
                }
            } catch (err) {
                toast('Ошибка: ' + err.message);
            }
        };

        // Кнопка «Отменить результат» — доступна, только если бой уже завершён
        const resetBtn = document.getElementById('reset-fight-result');
        const isCompleted = !!fight.completed && !!fight.winner_id;
        if (resetBtn) {
            resetBtn.style.display = isCompleted ? 'block' : 'none';
            resetBtn.onclick = async () => {
                if (!confirm('Сбросить результат этого боя?')) return;
                try {
                    await API.resetFightResult(this.currentTournamentId, this.currentFightId);
                    toast('✅ Результат боя сброшен', 'success');
                    document.getElementById('fight-modal').style.display = 'none';
                    const updatedTournament = await API.getTournament(this.currentTournamentId);
                    await this.renderBracket(updatedTournament);
                    await this.renderTournaments();
                } catch (err) {
                    toast('Ошибка: ' + err.message, 'error');
                }
            };
        }

        document.getElementById('fight-modal').style.display = 'block';
    },

async advanceWinner(tournamentId, fight, winnerId) {
        const tournament = await API.getTournament(tournamentId);
        const nextRound = fight.round + 1;

        // Все бои текущего раунда в той же весовой категории (по порядку id)
        const sameRoundFights = tournament.fights
            .filter(f => f.round === fight.round && f.weight_class === fight.weight_class)
            .sort((a, b) => a.id - b.id);
        const fightIndex = sameRoundFights.findIndex(f => f.id === fight.id);
        const targetIndex = Math.floor(fightIndex / 2);
        const isTopSlot = fightIndex % 2 === 0;

        // Бои следующего раунда в той же категории
        const nextRoundFights = tournament.fights
            .filter(f => f.round === nextRound && f.weight_class === fight.weight_class)
            .sort((a, b) => a.id - b.id);

        let nextFight = nextRoundFights[targetIndex];

        if (!nextFight) {
            // Создаём новый бой следующего раунда, победитель занимает своё место
            await API.addFight(tournamentId, {
                round: nextRound,
                athlete1_id: isTopSlot ? winnerId : null,
                athlete2_id: isTopSlot ? null : winnerId,
                weight_class: fight.weight_class
            });
            return;
        }

        // Победитель занимает слот в уже существующем бою
        const slotField = isTopSlot ? 'athlete1_id' : 'athlete2_id';
        if (nextFight[slotField] && nextFight[slotField] !== winnerId) {
            // Слот занят другим спортсменом — не затираем чужие данные
            return;
        }
        await API.updateFight(nextFight.id, { [slotField]: winnerId });
    },

    // ==================== РЕЗУЛЬТАТЫ ====================

    async renderResultsSelect() {
        const select = document.getElementById('results-tournament-select');
        const container = document.getElementById('results-container');

        try {
            const tournaments = await API.getTournaments();
            select.innerHTML = '<option value="">-- Выберите соревнование --</option>';
            tournaments.forEach(t => {
                const opt = document.createElement('option');
                opt.value = t.id;
                opt.textContent = `${t.name} (${this.fmtDate(t.date_from || t.date_to) || '—'})`;
                select.appendChild(opt);
            });

            // Загружаем сохраненное соревнование по умолчанию
            const defaultId = this.getDefaultTournamentId();
            if (defaultId) {
                const opt = select.querySelector(`option[value="${defaultId}"]`);
                if (opt) opt.selected = true;
            }

            const newSelect = select.cloneNode(true);
            select.parentNode.replaceChild(newSelect, select);

            // Устанавливаем значение после замены селекта
            if (defaultId) {
                newSelect.value = defaultId;
                newSelect.dispatchEvent(new Event('change'));
            }

            newSelect.addEventListener('change', async () => {
                const id = parseInt(newSelect.value);
                if (id) {
                    this.saveDefaultTournament(id);
                    const tournament = await API.getTournament(id);
                    
                    // Проверяем, есть ли участники
                    if (!tournament.athlete_ids || tournament.athlete_ids.length === 0) {
                        container.innerHTML = `
                            <div class="empty-state">
                                <p>⚠️ В соревновании нет участников!</p>
                                <p>Добавьте спортсменов через кнопку <strong>"👥 Участники"</strong> в разделе "Соревнования"</p>
                            </div>
                        `;
                        return;
                    }
                    
                    await this.renderResults(tournament);
                } else {
                    this.saveDefaultTournament(null);
                    container.innerHTML = '<div class="empty-state"><p>Выберите соревнование и сформируйте сетку для просмотра результатов</p></div>';
                }
            });

            // Загружаем результаты только если есть выбранный турнир
            if (defaultId) {
                const tournament = await API.getTournament(parseInt(defaultId));
                if (tournament.athlete_ids && tournament.athlete_ids.length > 0) {
                    await this.renderResults(tournament);
                } else {
                    container.innerHTML = '<div class="empty-state"><p>Выберите соревнование и сформируйте сетку для просмотра результатов</p></div>';
                }
            }
        } catch (err) {
            console.error('Ошибка:', err);
        }
    },

    async renderResults(tournament) {
        const container = document.getElementById('results-container');
        const athletes = await API.getAthletes();
        const idName = {};
        athletes.forEach(a => idName[a.id] = a);

        // Возвращает ID бойца из объекта слота или числового id
        const fid = (x) => (x && typeof x === 'object') ? x.athlete_id : x;

        // Результаты считаем из жеребьёвки + результатов пар (активный флоу),
        // т.к. таблица fights в этом флоу не заполняется.
        const [lottery, results] = await Promise.all([
            API.getLottery(tournament.id).catch(() => ({})),
            API.getBracketResults(tournament.id).catch(() => [])
        ]);

        const resultsMap = {};
        (Array.isArray(results) ? results : []).forEach(r => {
            resultsMap[`${r.weight_class}|${r.pair_index}`] = r;
        });

        if (!lottery || typeof lottery !== 'object' || Object.keys(lottery).length === 0) {
            container.innerHTML = '<div class="empty-state"><p>Проведите жеребьевку и сформируйте сетку, чтобы увидеть результаты</p></div>';
            return;
        }

        function getWeightLower(wc) {
            const m = wc.match(/(\d+(?:\.\d+)?)/);
            return m ? parseFloat(m[1]) : 999;
        }
        const sortedWeightClasses = Object.keys(lottery).sort((a, b) => {
            const gA = a.split(' ')[0], gB = b.split(' ')[0];
            const wA = getWeightLower(a), wB = getWeightLower(b);
            if (gA !== gB) return gA.localeCompare(gB);
            return wA - wB;
        });

// Количество дней: максимум из поля days и интервала дат (date_from..date_to)
        let daysFromDates = 1;
        if (tournament.date_from && tournament.date_to) {
            const diff = (new Date(tournament.date_to) - new Date(tournament.date_from)) / 86400000;
            if (isFinite(diff) && diff >= 0) daysFromDates = Math.floor(diff) + 1;
        }
        const tournamentDays = Math.max(daysFromDates, Math.max(1, parseInt(tournament.days) || 1));

        // Строит раунды из слотов, как в сетке. Возвращает { rounds, champion, maxRounds }
        function buildRounds(slots, wc) {
            const rounds = [];
            const firstRound = [];
            let pIdx = 0;
            for (let i = 0; i < slots.length - 1; i += 2) {
                firstRound.push({ a1: slots[i], a2: slots[i + 1], pairIndex: pIdx,
                    result: resultsMap[`${wc}|${pIdx}`] || null });
                pIdx++;
            }
            if (slots.length % 2 !== 0) {
                firstRound.push({ a1: slots[slots.length - 1], a2: null, pairIndex: null,
                    result: { winner_id: slots[slots.length - 1].athlete_id, method: 'авто' } });
            }
            rounds.push(firstRound);

            let roundNum = 1;
            while (true) {
                const prev = rounds[rounds.length - 1];
                const winners = prev.map(f => f.result ? f.result.winner_id : null);
                if (winners.some(w => w === null)) break;
                if (winners.length <= 1) break;
                roundNum++;
                const offset = (roundNum - 1) * 1000;
                const next = [];
                for (let i = 0; i < winners.length - 1; i += 2) {
                    const pi = offset + i;
                    next.push({ a1: winners[i], a2: winners[i + 1], pairIndex: pi,
                        result: resultsMap[`${wc}|${pi}`] || null });
                }
                if (next.length === 0) break;
                rounds.push(next);
            }

            // Чемпион — единственный победитель верхнего раунда, если все пары решены
            let champion = null;
            const top = rounds[rounds.length - 1];
            const topWinners = top.map(f => f.result ? f.result.winner_id : null);
            const anyUnresolved = rounds.some(r => r.some(f => f.a2 !== null && !f.result));
            if (!anyUnresolved && topWinners.length === 1 && topWinners[0] != null) {
                champion = topWinners[0];
            }
            return { rounds, champion, maxRounds: rounds.length };
        }

        let html = '';
        let hasResults = false;

        for (const wc of sortedWeightClasses) {
            const slots = [...(lottery[wc] || [])].sort((a, b) => a.slot - b.slot);
            if (slots.length === 0) continue;

            const { rounds, champion, maxRounds } = buildRounds(slots, wc);

            // Категория с ещё не решёнными боями финалист не даёт
            if (champion == null) continue;

            hasResults = true;
            const gold = idName[champion] ? { name: idName[champion].name, id: champion } : null;

            // Серебро — проигравший финала; бронза — ОБА проигравших полуфинала (две бронзы)
            let silver = null;
            const bronzes = [];   // два бронзовых призёра
            const finalRound = rounds[rounds.length - 1];
            if (finalRound.length === 1 && finalRound[0].a2 !== null) {
                const finalPair = finalRound[0];
                const la = fid(finalPair.a1), lb = fid(finalPair.a2);
                const loserId = (la !== champion) ? la : lb;
                if (loserId != null && idName[loserId]) silver = { name: idName[loserId].name, id: loserId };
            }
            if (maxRounds >= 2) {
                const semiRound = rounds[rounds.length - 2];
                for (const sp of semiRound) {
                    if (!sp.result || sp.a2 === null) continue;
                    const la = fid(sp.a1), lb = fid(sp.a2);
                    const loserId = (sp.result.winner_id === la) ? lb : la;
                    if (loserId != null && idName[loserId] && loserId !== champion) {
                        if (bronzes.length < 2) {
                            bronzes.push({ name: idName[loserId].name, id: loserId });
                        }
                    }
                }
            }

            // День финала
            const finalDay = this.roundToDay(maxRounds, maxRounds, tournamentDays);
            const finalName = finalRound.length === 1 && finalRound[0].a2 !== null ? 'Финал' : 'Победитель без соперника';

            html += `<div class="bracket-round"><h3 class="weight-title">${this.escapeHtml(wc)} <span class="day-tag">День ${finalDay}</span></h3>`;

            const medalStyle = (color, bg) => `style="border-color:${color}; background:${bg};"`;
            // Победитель категории (Золото)
            html += `<div class="result-card" ${medalStyle('#ffd700', 'rgba(255,215,0,0.1)')}>
                <div class="result-fighters" style="justify-content:center; gap:16px;">
                    <span style="font-size:1.8em;">🥇</span>
                    <span style="font-size:1.3em; font-weight:700; color:#ffd700;">${gold ? this.escapeHtml(gold.name) : '—'}</span>
                </div>
            </div>`;

            if (silver) {
                html += `<div class="result-card" ${medalStyle('#c0c0c0', 'rgba(192,192,192,0.08)')}>
                    <div class="result-fighters" style="justify-content:center; gap:16px;">
                        <span style="font-size:1.6em;">🥈</span>
                        <span style="font-size:1.15em; font-weight:600; color:#c0c0c0;">${this.escapeHtml(silver.name)}</span>
                    </div>
                </div>`;
            }
            // Две бронзы — оба проигравших полуфиналиста
            bronzes.forEach(b => {
                html += `<div class="result-card" ${medalStyle('#cd7f32', 'rgba(205,127,50,0.08)')}>
                    <div class="result-fighters" style="justify-content:center; gap:16px;">
                        <span style="font-size:1.6em;">🥉</span>
                        <span style="font-size:1.15em; font-weight:600; color:#cd7f32;">${this.escapeHtml(b.name)}</span>
                    </div>
                </div>`;
            });

            html += '</div>';
        }

        if (!hasResults) {
            container.innerHTML = '<div class="empty-state"><p>В этом соревновании ещё нет завершённых категорий. Определите победителей всех боёв категории, чтобы они появились здесь.</p></div>';
            return;
        }

        container.innerHTML = html;
    },

    // ==================== ВСПОМОГАТЕЛЬНЫЕ ====================

    // Дата YYYY-MM-DD -> ДД.ММ.ГГГГ
    fmtDate(s) {
        if (!s) return s;
        const str = String(s);
        const m = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
        return m ? `${m[3]}.${m[2]}.${m[1]}` : str;
    },

    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    },

    async renderAll() {
        await this.renderTournaments();
        this.setupLottery();
        // Не вызываем renderBracketSelect и renderResultsSelect, чтобы не очищать контейнеры
        // Данные загрузятся из БД в loadBracketFromDB и loadLotteryFromDB
        this.setupWeighin();
    }
};

// ======================== ЗАПУСК ========================

document.addEventListener('DOMContentLoaded', () => {
    App.init();
});