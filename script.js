/**
 * FinPulse - script.js
 *
 * APIs used:
 *  1. CoinGecko API       - Live crypto market data (data API)
 *  2. Geolocation API     - User location for market status (browser API)
 *  3. Web Speech API      - Voice market briefing (browser API)
 *  4. Intersection Observer API - Scroll-triggered animations (browser API)
 */

'use strict';

/* ===================================================
   CONFIG
   =================================================== */
const CONFIG = {
  COINGECKO_BASE:   'https://api.coingecko.com/api/v3',
  COINS_COUNT:      25,
  REFRESH_INTERVAL: 60000, // ms
};

/* ===================================================
   STATE
   =================================================== */
const state = {
  coins:          [],
  globalData:     null,
  watchlist:      JSON.parse(localStorage.getItem('finpulse-watchlist') || '[]'),
  currentSort:    'market_cap',
  searchQuery:    '',
  isSpeaking:     false,
  refreshTimer:   null,
  countdownVal:   60,
  countdownTimer: null,
};

/* ===================================================
   DOM REFERENCES
   =================================================== */
const $ = (id) => document.getElementById(id);
const $$ = (sel) => document.querySelectorAll(sel);

const DOM = {
  body:              document.body,
  themeToggle:       $('theme-toggle'),
  tickerTrack:       $('ticker-track'),

  // Hero stats
  totalMarketCap:    $('total-market-cap'),
  totalVolume:       $('total-volume'),
  btcDominance:      $('btc-dominance'),
  activeCoins:       $('active-coins'),

  // Market status
  locationPanel:     $('location-panel'),
  locationPrompt:    $('location-prompt'),
  getLocationBtn:    $('get-location-btn'),
  skipLocationBtn:   $('skip-location-btn'),
  exchangesGrid:     $('exchanges-grid'),

  // Crypto dashboard
  cryptoSearch:      $('crypto-search'),
  sortBtns:          $$('.sort-btn'),
  refreshText:       $('refresh-text'),
  sentimentFill:     $('sentiment-fill'),
  sentimentScore:    $('sentiment-score'),
  cryptoLoading:     $('crypto-loading'),
  cryptoError:       $('crypto-error'),
  cryptoErrorMsg:    $('crypto-error-msg'),
  retryBtn:          $('retry-btn'),
  cryptoTableWrap:   $('crypto-table-wrapper'),
  cryptoTbody:       $('crypto-tbody'),

  // Voice
  voiceCard:         $('voice-card'),
  voiceWaveform:     $('voice-waveform'),
  voiceSelect:       $('voice-select'),
  voiceRate:         $('voice-rate'),
  voiceRateDisplay:  $('voice-rate-display'),
  speakBtn:          $('speak-btn'),
  stopBtn:           $('stop-btn'),
  voiceError:        $('voice-error'),
  voiceTranscriptWrap: $('voice-transcript-wrap'),
  voiceTranscript:   $('voice-transcript'),

  // Watchlist
  watchlistEmpty:    $('watchlist-empty'),
  watchlistGrid:     $('watchlist-grid'),

  // Newsletter
  newsletterForm:    $('newsletter-form'),
  formMessage:       $('form-message'),
};

/* ===================================================
   UTILITIES
   =================================================== */

function formatPrice(price) {
  if (price == null) return 'N/A';
  const opts =
    price >= 1
      ? { style: 'currency', currency: 'USD', minimumFractionDigits: 2, maximumFractionDigits: 2 }
      : { style: 'currency', currency: 'USD', minimumFractionDigits: 4, maximumFractionDigits: 6 };
  return new Intl.NumberFormat('en-US', opts).format(price);
}

function formatLargeNum(value) {
  if (value == null) return 'N/A';
  if (value >= 1e12) return `$${(value / 1e12).toFixed(2)}T`;
  if (value >= 1e9)  return `$${(value / 1e9).toFixed(2)}B`;
  if (value >= 1e6)  return `$${(value / 1e6).toFixed(2)}M`;
  return `$${value.toLocaleString()}`;
}

function formatChange(pct) {
  if (pct == null) return 'N/A';
  const sign = pct >= 0 ? '+' : '';
  return `${sign}${pct.toFixed(2)}%`;
}

function changeClass(pct) {
  if (pct == null) return '';
  return pct >= 0 ? 'positive' : 'negative';
}

function escHtml(str) {
  const d = document.createElement('div');
  d.textContent = str;
  return d.innerHTML;
}

/* ===================================================
   THEME TOGGLE
   =================================================== */

function initTheme() {
  const saved = localStorage.getItem('finpulse-theme');
  if (saved === 'light') {
    DOM.body.classList.add('light-mode');
    DOM.themeToggle.querySelector('.theme-icon').textContent = '\u2600'; // sun
  }
}

function handleThemeToggle() {
  DOM.body.classList.toggle('light-mode');
  const isLight = DOM.body.classList.contains('light-mode');
  DOM.themeToggle.querySelector('.theme-icon').textContent = isLight ? '\u2600' : '\u263D';
  localStorage.setItem('finpulse-theme', isLight ? 'light' : 'dark');
}

/* ===================================================
   COINGECKO API - GLOBAL DATA
   =================================================== */

async function fetchGlobalData() {
  try {
    const res = await fetch(`${CONFIG.COINGECKO_BASE}/global`);
    if (!res.ok) return;
    const json = await res.json();
    state.globalData = json.data;
    renderHeroStats();
  } catch {
    // Non-fatal; hero stats remain as skeleton
  }
}

function renderHeroStats() {
  const gd = state.globalData;
  if (!gd) return;

  DOM.totalMarketCap.textContent = formatLargeNum(gd.total_market_cap?.usd);
  DOM.totalVolume.textContent    = formatLargeNum(gd.total_volume?.usd);
  DOM.btcDominance.textContent   = gd.market_cap_percentage?.btc
    ? `${gd.market_cap_percentage.btc.toFixed(1)}%`
    : 'N/A';
  DOM.activeCoins.textContent    = gd.active_cryptocurrencies
    ? gd.active_cryptocurrencies.toLocaleString()
    : 'N/A';
}

/* ===================================================
   COINGECKO API - COINS
   =================================================== */

async function fetchCoins() {
  showCryptoLoading();

  try {
    const url =
      `${CONFIG.COINGECKO_BASE}/coins/markets` +
      `?vs_currency=usd` +
      `&order=market_cap_desc` +
      `&per_page=${CONFIG.COINS_COUNT}` +
      `&page=1` +
      `&sparkline=false` +
      `&price_change_percentage=24h,7d`;

    const res = await fetch(url);

    if (res.status === 429) {
      throw new Error('Rate limit reached. Please wait a moment before refreshing.');
    }
    if (!res.ok) {
      throw new Error(`CoinGecko returned an error (${res.status}). Please try again.`);
    }

    const data = await res.json();
    if (!Array.isArray(data) || data.length === 0) {
      throw new Error('No market data returned. Please try again shortly.');
    }

    state.coins = data;
    renderCryptoTable();
    buildTicker();
    updateSentimentBar();
    renderWatchlist();
    startCountdown();
    showCryptoTable();

  } catch (err) {
    showCryptoError(err.message);
  }
}

/* ---- Crypto table display helpers ---- */

function showCryptoLoading() {
  DOM.cryptoLoading.style.display  = 'flex';
  DOM.cryptoError.style.display    = 'none';
  DOM.cryptoTableWrap.style.display = 'none';
}

function showCryptoTable() {
  DOM.cryptoLoading.style.display   = 'none';
  DOM.cryptoError.style.display     = 'none';
  DOM.cryptoTableWrap.style.display = 'block';
}

function showCryptoError(msg) {
  DOM.cryptoLoading.style.display   = 'none';
  DOM.cryptoTableWrap.style.display = 'none';
  DOM.cryptoErrorMsg.textContent    = `\u26A0 ${msg}`;
  DOM.cryptoError.style.display     = 'flex';
}

/* ---- Sorting / Filtering ---- */

function getSortedCoins() {
  let coins = [...state.coins];

  if (state.searchQuery) {
    const q = state.searchQuery.toLowerCase();
    coins = coins.filter(
      (c) => c.name.toLowerCase().includes(q) || c.symbol.toLowerCase().includes(q)
    );
  }

  switch (state.currentSort) {
    case 'price':
      coins.sort((a, b) => b.current_price - a.current_price);
      break;
    case 'change_24h':
      coins.sort(
        (a, b) => (b.price_change_percentage_24h || 0) - (a.price_change_percentage_24h || 0)
      );
      break;
    case 'volume':
      coins.sort((a, b) => b.total_volume - a.total_volume);
      break;
    default:
      // market_cap - already sorted from API
      break;
  }

  return coins;
}

/* ---- Render table ---- */

function renderCryptoTable() {
  const coins = getSortedCoins();

  if (coins.length === 0) {
    DOM.cryptoTbody.innerHTML =
      '<tr class="no-results"><td colspan="8">No coins match your search.</td></tr>';
    return;
  }

  DOM.cryptoTbody.innerHTML = coins
    .map((coin, idx) => buildCoinRow(coin, idx + 1))
    .join('');

  // Attach watchlist button listeners
  DOM.cryptoTbody.querySelectorAll('.watchlist-star-btn').forEach((btn) => {
    btn.addEventListener('click', () => toggleWatchlist(btn.dataset.id));
  });
}

function buildCoinRow(coin, rank) {
  const change24  = coin.price_change_percentage_24h;
  const change7d  = coin.price_change_percentage_7d_in_currency;
  const starred   = state.watchlist.includes(coin.id);

  return `
    <tr>
      <td class="td-rank">${rank}</td>
      <td>
        <div class="coin-info-cell">
          <img src="${escHtml(coin.image)}" alt="${escHtml(coin.name)} logo" width="26" height="26" loading="lazy" />
          <div class="coin-name-wrap">
            <span class="coin-name">${escHtml(coin.name)}</span>
            <span class="coin-symbol">${escHtml(coin.symbol.toUpperCase())}</span>
          </div>
        </div>
      </td>
      <td class="td-price td-mono">${formatPrice(coin.current_price)}</td>
      <td class="${changeClass(change24)} td-mono">${formatChange(change24)}</td>
      <td class="${changeClass(change7d)} td-mono">${formatChange(change7d)}</td>
      <td class="td-mono">${formatLargeNum(coin.market_cap)}</td>
      <td class="td-mono">${formatLargeNum(coin.total_volume)}</td>
      <td>
        <button
          class="watchlist-star-btn ${starred ? 'starred' : ''}"
          data-id="${escHtml(coin.id)}"
          aria-label="${starred ? 'Remove' : 'Add'} ${escHtml(coin.name)} ${starred ? 'from' : 'to'} watchlist"
          aria-pressed="${starred}"
        >
          ${starred ? '\u2605' : '\u2606'}
        </button>
      </td>
    </tr>
  `;
}

/* ---- Sentiment bar ---- */

function updateSentimentBar() {
  const total = state.coins.length;
  if (total === 0) return;
  const up = state.coins.filter((c) => (c.price_change_percentage_24h || 0) >= 0).length;
  const pct = Math.round((up / total) * 100);

  DOM.sentimentFill.style.width = `${pct}%`;
  DOM.sentimentFill.setAttribute('aria-label', `Market sentiment: ${pct}% bullish`);
  DOM.sentimentScore.textContent = `${pct}% \u2191`;
}

/* ---- Countdown / refresh ---- */

function startCountdown() {
  state.countdownVal = 60;
  const now = new Date().toLocaleTimeString();
  DOM.refreshText.textContent = `Updated ${now} - next in ${state.countdownVal}s`;

  if (state.countdownTimer) clearInterval(state.countdownTimer);
  state.countdownTimer = setInterval(() => {
    state.countdownVal -= 1;
    if (state.countdownVal <= 0) {
      DOM.refreshText.textContent = 'Refreshing...';
    } else {
      const t = new Date().toLocaleTimeString();
      DOM.refreshText.textContent = `Updated ${t} - next in ${state.countdownVal}s`;
    }
  }, 1000);
}

/* ===================================================
   TICKER (uses fetched coin data)
   =================================================== */

function buildTicker() {
  if (state.coins.length === 0) return;

  const items = [...state.coins, ...state.coins]
    .map((c) => {
      const pct     = c.price_change_percentage_24h;
      const cls     = pct != null && pct >= 0 ? 'ticker-up' : 'ticker-down';
      const sign    = pct != null && pct >= 0 ? '+' : '';
      const pctStr  = pct != null ? `${sign}${pct.toFixed(2)}%` : 'N/A';
      return `<span class="ticker-item">
        <strong>${escHtml(c.symbol.toUpperCase())}</strong>
        ${formatPrice(c.current_price)}
        <span class="${cls}">${pctStr}</span>
      </span>`;
    })
    .join('');

  DOM.tickerTrack.innerHTML = items;
}

/* ===================================================
   GEOLOCATION API
   =================================================== */

const EXCHANGES = [
  {
    name:      'NYSE / NASDAQ',
    location:  'New York, USA',
    flag:      '\uD83C\uDDFA\uD83C\uDDF8',
    timezone:  'America/New_York',
    openHour:  9,  openMin:  30,
    closeHour: 16, closeMin: 0,
  },
  {
    name:      'London Stock Exchange',
    location:  'London, UK',
    flag:      '\uD83C\uDDEC\uD83C\uDDE7',
    timezone:  'Europe/London',
    openHour:  8,  openMin:  0,
    closeHour: 16, closeMin: 30,
  },
  {
    name:      'Euronext',
    location:  'Amsterdam / Paris',
    flag:      '\uD83C\uDDEA\uD83C\uDDFA',
    timezone:  'Europe/Paris',
    openHour:  9,  openMin:  0,
    closeHour: 17, closeMin: 30,
  },
  {
    name:      'Tokyo Stock Exchange',
    location:  'Tokyo, Japan',
    flag:      '\uD83C\uDDEF\uD83C\uDDF5',
    timezone:  'Asia/Tokyo',
    openHour:  9,  openMin:  0,
    closeHour: 15, closeMin: 30,
  },
  {
    name:      'Shanghai Stock Exchange',
    location:  'Shanghai, China',
    flag:      '\uD83C\uDDE8\uD83C\uDDF3',
    timezone:  'Asia/Shanghai',
    openHour:  9,  openMin:  30,
    closeHour: 15, closeMin: 0,
  },
  {
    name:      'BSE / NSE',
    location:  'Mumbai, India',
    flag:      '\uD83C\uDDEE\uD83C\uDDF3',
    timezone:  'Asia/Kolkata',
    openHour:  9,  openMin:  15,
    closeHour: 15, closeMin: 30,
  },
];

/**
 * Determine if an exchange is currently open by converting the current moment
 * into the exchange's local timezone and checking against open/close hours.
 */
function isExchangeOpen(ex) {
  const now = new Date();
  // Get the current time in the exchange's timezone as a locale string
  const localStr = now.toLocaleString('en-US', { timeZone: ex.timezone, hour12: false });
  // localStr looks like "3/22/2026, 09:45:00"
  const parts = localStr.split(', ');
  if (parts.length < 2) return false;

  const timeParts = parts[1].split(':');
  const hour = parseInt(timeParts[0], 10);
  const min  = parseInt(timeParts[1], 10);

  // Check weekday in the exchange timezone
  const dayStr = now.toLocaleDateString('en-US', { timeZone: ex.timezone, weekday: 'short' });
  const isWeekday = !['Sat', 'Sun'].includes(dayStr);
  if (!isWeekday) return false;

  const nowMins   = hour * 60 + min;
  const openMins  = ex.openHour  * 60 + ex.openMin;
  const closeMins = ex.closeHour * 60 + ex.closeMin;

  return nowMins >= openMins && nowMins < closeMins;
}

function getExchangeLocalTime(ex) {
  const now = new Date();
  return now.toLocaleTimeString('en-US', {
    timeZone: ex.timezone,
    hour:     '2-digit',
    minute:   '2-digit',
    hour12:   true,
  });
}

function renderExchanges() {
  DOM.exchangesGrid.innerHTML = EXCHANGES.map((ex) => {
    const open      = isExchangeOpen(ex);
    const localTime = getExchangeLocalTime(ex);
    const openTime  = `${String(ex.openHour).padStart(2,'0')}:${String(ex.openMin).padStart(2,'0')}`;
    const closeTime = `${String(ex.closeHour).padStart(2,'0')}:${String(ex.closeMin).padStart(2,'0')}`;

    return `
      <div class="exchange-card ${open ? 'is-open' : 'is-closed'}">
        <div class="exchange-flag" aria-hidden="true">${ex.flag}</div>
        <div class="exchange-info">
          <h3>${escHtml(ex.name)}</h3>
          <p class="exchange-location">${escHtml(ex.location)}</p>
          <p class="exchange-time">Local: ${localTime} | Hours: ${openTime}-${closeTime}</p>
        </div>
        <div>
          <span class="status-badge ${open ? 'status-open' : 'status-closed'}">
            ${open ? 'OPEN' : 'CLOSED'}
          </span>
        </div>
      </div>
    `;
  }).join('');

  // Trigger reveal animations for exchange cards
  requestAnimationFrame(() => {
    DOM.exchangesGrid.querySelectorAll('.exchange-card').forEach((el, i) => {
      setTimeout(() => el.classList.add('revealed'), i * 80);
    });
  });
}

function handleGeolocationSuccess(pos) {
  const { latitude, longitude } = pos.coords;
  const tz = Intl.DateTimeFormat().resolvedOptions().timeZone;

  function updateLocalTime() {
    return new Date().toLocaleTimeString('en-US', {
      timeZone: tz,
      hour:     '2-digit',
      minute:   '2-digit',
      second:   '2-digit',
      hour12:   true,
    });
  }

  DOM.locationPanel.innerHTML = `
    <div class="location-found">
      <div class="location-icon" aria-hidden="true">\uD83D\uDCCD</div>
      <div class="location-details">
        <p><strong id="local-clock">${updateLocalTime()}</strong></p>
        <p>Timezone: <strong>${escHtml(tz)}</strong></p>
        <p>Coordinates: <strong>${latitude.toFixed(4)}, ${longitude.toFixed(4)}</strong></p>
      </div>
    </div>
  `;

  // Live clock
  setInterval(() => {
    const cl = document.getElementById('local-clock');
    if (cl) cl.textContent = updateLocalTime();
    renderExchanges(); // re-evaluate open/closed status
  }, 10000);
}

function handleGeolocationError(err) {
  let msg = 'Could not retrieve your location.';
  if (err.code === err.PERMISSION_DENIED) {
    msg = 'Location access was denied. Exchange hours are still shown below based on UTC.';
  } else if (err.code === err.POSITION_UNAVAILABLE) {
    msg = 'Location data is unavailable. Exchange hours shown below.';
  } else if (err.code === err.TIMEOUT) {
    msg = 'Location request timed out. Exchange hours shown below.';
  }

  DOM.locationPanel.innerHTML = `
    <div class="location-error">
      <span aria-hidden="true">\u26A0</span>
      ${escHtml(msg)}
    </div>
  `;
}

function requestGeolocation() {
  if (!navigator.geolocation) {
    DOM.locationPanel.innerHTML = `
      <div class="location-error">
        <span aria-hidden="true">\u26A0</span>
        Geolocation is not supported by your browser. Exchange hours shown below.
      </div>
    `;
    return;
  }

  DOM.locationPanel.innerHTML = `
    <div class="location-loading">
      <div class="spinner sm" role="status"><span class="sr-only">Locating you...</span></div>
      <p>Getting your location...</p>
    </div>
  `;

  navigator.geolocation.getCurrentPosition(
    handleGeolocationSuccess,
    handleGeolocationError,
    { timeout: 10000, maximumAge: 600000 }
  );
}

function skipGeolocation() {
  DOM.locationPanel.innerHTML = `
    <div class="location-error">
      <span aria-hidden="true">\uD83D\uDCCD</span>
      No location data used. Exchange hours are shown based on current UTC time.
    </div>
  `;
}

/* ===================================================
   WEB SPEECH API
   =================================================== */

function initSpeechSynthesis() {
  if (!('speechSynthesis' in window)) {
    DOM.voiceError.textContent =
      '\u26A0 Text-to-speech is not supported in your browser. Try Chrome, Edge, or Safari.';
    DOM.voiceError.style.display = 'block';
    DOM.speakBtn.disabled = true;
    DOM.speakBtn.setAttribute('aria-disabled', 'true');
    return;
  }

  function loadVoices() {
    const voices = window.speechSynthesis.getVoices().filter((v) => v.lang.startsWith('en'));
    DOM.voiceSelect.innerHTML = voices.length
      ? voices.map((v) => `<option value="${escHtml(v.name)}">${escHtml(v.name)} (${v.lang})</option>`).join('')
      : '<option value="">Default Voice</option>';
  }

  loadVoices();
  window.speechSynthesis.onvoiceschanged = loadVoices;
}

function buildBriefingText() {
  if (state.coins.length === 0) {
    return 'No market data is currently available. Please wait for the dashboard to load and try again.';
  }

  const btc = state.coins.find((c) => c.id === 'bitcoin');
  const eth = state.coins.find((c) => c.id === 'ethereum');

  const sorted24h = [...state.coins].sort(
    (a, b) => (b.price_change_percentage_24h || 0) - (a.price_change_percentage_24h || 0)
  );
  const gainer = sorted24h[0];
  const loser  = sorted24h[sorted24h.length - 1];

  let text = 'Welcome to your FinPulse voice market briefing. ';

  if (state.globalData) {
    const cap = formatLargeNum(state.globalData.total_market_cap?.usd);
    const btcDom = state.globalData.market_cap_percentage?.btc?.toFixed(1);
    text += `The total cryptocurrency market cap is ${cap}. Bitcoin dominance is ${btcDom} percent. `;
  }

  if (btc) {
    const dir   = (btc.price_change_percentage_24h || 0) >= 0 ? 'up' : 'down';
    const delta = Math.abs(btc.price_change_percentage_24h || 0).toFixed(2);
    text += `Bitcoin is trading at ${formatPrice(btc.current_price)}, ${dir} ${delta} percent in the last 24 hours. `;
  }

  if (eth) {
    const dir   = (eth.price_change_percentage_24h || 0) >= 0 ? 'up' : 'down';
    const delta = Math.abs(eth.price_change_percentage_24h || 0).toFixed(2);
    text += `Ethereum is at ${formatPrice(eth.current_price)}, ${dir} ${delta} percent. `;
  }

  if (gainer) {
    text += `Today's biggest gainer is ${gainer.name}, up ${(gainer.price_change_percentage_24h || 0).toFixed(2)} percent. `;
  }

  if (loser) {
    text += `The biggest decliner is ${loser.name}, down ${Math.abs(loser.price_change_percentage_24h || 0).toFixed(2)} percent. `;
  }

  const upCount = state.coins.filter((c) => (c.price_change_percentage_24h || 0) >= 0).length;
  const pct     = Math.round((upCount / state.coins.length) * 100);
  text += `Overall, ${upCount} of the top ${state.coins.length} coins are in positive territory, giving a bullish sentiment of ${pct} percent. `;

  text += 'As always, this is for informational purposes only. Please do your own research before making any investment decisions. ';
  text += 'This has been your FinPulse voice market briefing. Good luck out there.';

  return text;
}

function speakBriefing() {
  if (!('speechSynthesis' in window)) return;

  window.speechSynthesis.cancel(); // cancel any in-progress speech

  const text       = buildBriefingText();
  const utterance  = new SpeechSynthesisUtterance(text);

  // Voice selection
  const selectedName = DOM.voiceSelect.value;
  if (selectedName) {
    const voice = window.speechSynthesis.getVoices().find((v) => v.name === selectedName);
    if (voice) utterance.voice = voice;
  }

  utterance.rate   = parseFloat(DOM.voiceRate.value) || 1;
  utterance.pitch  = 1;
  utterance.volume = 1;

  utterance.onstart = () => {
    state.isSpeaking = true;
    DOM.speakBtn.style.display  = 'none';
    DOM.stopBtn.style.display   = 'inline-flex';
    DOM.voiceWaveform.classList.add('active');
    DOM.voiceTranscript.textContent    = text;
    DOM.voiceTranscriptWrap.style.display = 'block';
    DOM.voiceError.style.display       = 'none';
  };

  utterance.onend = () => {
    state.isSpeaking = false;
    DOM.speakBtn.style.display  = 'inline-flex';
    DOM.stopBtn.style.display   = 'none';
    DOM.voiceWaveform.classList.remove('active');
  };

  utterance.onerror = (e) => {
    state.isSpeaking = false;
    DOM.speakBtn.style.display  = 'inline-flex';
    DOM.stopBtn.style.display   = 'none';
    DOM.voiceWaveform.classList.remove('active');

    if (e.error !== 'interrupted') {
      DOM.voiceError.textContent  = `\u26A0 Speech error: "${e.error}". Try a different voice or browser.`;
      DOM.voiceError.style.display = 'block';
    }
  };

  window.speechSynthesis.speak(utterance);
}

function stopSpeaking() {
  if ('speechSynthesis' in window) {
    window.speechSynthesis.cancel();
  }
  state.isSpeaking = false;
  DOM.speakBtn.style.display  = 'inline-flex';
  DOM.stopBtn.style.display   = 'none';
  DOM.voiceWaveform.classList.remove('active');
}

/* ===================================================
   WATCHLIST
   =================================================== */

function toggleWatchlist(coinId) {
  const idx = state.watchlist.indexOf(coinId);
  if (idx === -1) {
    state.watchlist.push(coinId);
  } else {
    state.watchlist.splice(idx, 1);
  }
  localStorage.setItem('finpulse-watchlist', JSON.stringify(state.watchlist));
  renderCryptoTable();
  renderWatchlist();
}

function renderWatchlist() {
  const coins = state.coins.filter((c) => state.watchlist.includes(c.id));

  if (coins.length === 0) {
    DOM.watchlistEmpty.style.display = 'block';
    DOM.watchlistGrid.innerHTML = '';
    return;
  }

  DOM.watchlistEmpty.style.display = 'none';
  DOM.watchlistGrid.innerHTML = coins.map((coin) => {
    const change = coin.price_change_percentage_24h;
    return `
      <div class="wl-card">
        <div class="wl-header">
          <img src="${escHtml(coin.image)}" alt="${escHtml(coin.name)}" width="28" height="28" loading="lazy" />
          <div class="coin-name-wrap">
            <span class="coin-name">${escHtml(coin.name)}</span>
            <span class="coin-symbol">${escHtml(coin.symbol.toUpperCase())}</span>
          </div>
          <button
            class="wl-remove-btn"
            data-id="${escHtml(coin.id)}"
            aria-label="Remove ${escHtml(coin.name)} from watchlist"
          >
            \u00D7
          </button>
        </div>
        <div class="wl-price">${formatPrice(coin.current_price)}</div>
        <div class="wl-change ${changeClass(change)}">${formatChange(change)} (24h)</div>
        <div class="wl-cap">Cap: ${formatLargeNum(coin.market_cap)}</div>
      </div>
    `;
  }).join('');

  // Listeners for remove buttons
  DOM.watchlistGrid.querySelectorAll('.wl-remove-btn').forEach((btn) => {
    btn.addEventListener('click', () => toggleWatchlist(btn.dataset.id));
  });

  // Animate watchlist cards
  requestAnimationFrame(() => {
    DOM.watchlistGrid.querySelectorAll('.wl-card').forEach((el, i) => {
      setTimeout(() => el.classList.add('revealed'), i * 60);
    });
  });
}

/* ===================================================
   NEWSLETTER FORM
   =================================================== */

function setFieldValidity(inputId, errorId, valid, message) {
  const input = document.getElementById(inputId);
  const error = document.getElementById(errorId);
  if (!input || !error) return valid;

  error.textContent = valid ? '' : message;
  if (valid) {
    input.removeAttribute('aria-invalid');
  } else {
    input.setAttribute('aria-invalid', 'true');
  }
  return valid;
}

function validateNewsletter() {
  const firstName = document.getElementById('first-name').value.trim();
  const lastName  = document.getElementById('last-name').value.trim();
  const email     = document.getElementById('email').value.trim();
  const interest  = document.getElementById('interest').value;
  const terms     = document.getElementById('terms').checked;

  const emailRx = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;

  const results = [
    setFieldValidity('first-name', 'first-name-error', firstName.length >= 2, 'Please enter your first name (at least 2 characters).'),
    setFieldValidity('last-name',  'last-name-error',  lastName.length >= 2,  'Please enter your last name (at least 2 characters).'),
    setFieldValidity('email',      'email-error',       emailRx.test(email),   'Please enter a valid email address.'),
    setFieldValidity('interest',   'interest-error',    interest !== '',        'Please select your primary focus area.'),
    setFieldValidity('terms',      'terms-error',        terms,                 'You must agree to receive emails to subscribe.'),
  ];

  return results.every(Boolean);
}

function handleNewsletterSubmit(e) {
  e.preventDefault();
  if (!validateNewsletter()) return;

  const firstName = document.getElementById('first-name').value.trim();

  DOM.formMessage.className           = 'form-message success';
  DOM.formMessage.textContent         = `You're in, ${firstName}! Welcome to FinPulse. Watch your inbox for your first briefing.`;
  DOM.formMessage.style.display       = 'block';

  DOM.newsletterForm.reset();

  setTimeout(() => {
    DOM.formMessage.style.display = 'none';
  }, 8000);
}

/* ===================================================
   INTERSECTION OBSERVER API - Scroll animations
   =================================================== */

function initScrollAnimations() {
  if (!('IntersectionObserver' in window)) {
    // Graceful fallback: just show everything immediately
    $$('.animate-reveal').forEach((el) => el.classList.add('revealed'));
    return;
  }

  const observer = new IntersectionObserver(
    (entries) => {
      entries.forEach((entry) => {
        if (entry.isIntersecting) {
          entry.target.classList.add('revealed');
          observer.unobserve(entry.target); // animate once
        }
      });
    },
    { threshold: 0.12, rootMargin: '0px 0px -40px 0px' }
  );

  $$('.animate-reveal').forEach((el) => observer.observe(el));

  // Also observe stat cards after they load
  $$('.stat-card').forEach((el, i) => {
    // Hero stat cards need a slight stagger
    const cardObserver = new IntersectionObserver(
      (entries) => {
        entries.forEach((entry) => {
          if (entry.isIntersecting) {
            setTimeout(() => entry.target.classList.add('revealed'), i * 100);
            cardObserver.unobserve(entry.target);
          }
        });
      },
      { threshold: 0.1 }
    );
    cardObserver.observe(el);
  });
}

/* ===================================================
   SMOOTH SCROLL
   =================================================== */

function initSmoothScroll() {
  $$('a[href^="#"]').forEach((anchor) => {
    anchor.addEventListener('click', function (e) {
      const target = document.querySelector(this.getAttribute('href'));
      if (!target) return;
      e.preventDefault();
      target.scrollIntoView({ behavior: 'smooth', block: 'start' });
    });
  });
}

/* ===================================================
   AUTO-REFRESH
   =================================================== */

function startAutoRefresh() {
  if (state.refreshTimer) clearInterval(state.refreshTimer);
  state.refreshTimer = setInterval(() => {
    fetchCoins();
    fetchGlobalData();
    renderExchanges(); // re-check open/closed
  }, CONFIG.REFRESH_INTERVAL);
}

/* ===================================================
   INITIALISE
   =================================================== */

async function init() {
  // Theme
  initTheme();
  DOM.themeToggle.addEventListener('click', handleThemeToggle);

  // Smooth scroll
  initSmoothScroll();

  // Intersection Observer
  initScrollAnimations();

  // Web Speech API
  initSpeechSynthesis();

  // Geolocation buttons
  if (DOM.getLocationBtn) {
    DOM.getLocationBtn.addEventListener('click', requestGeolocation);
  }
  if (DOM.skipLocationBtn) {
    DOM.skipLocationBtn.addEventListener('click', skipGeolocation);
  }

  // Initial exchange render (without location)
  renderExchanges();

  // Crypto search
  DOM.cryptoSearch.addEventListener('input', (e) => {
    state.searchQuery = e.target.value;
    renderCryptoTable();
  });

  // Sort buttons
  DOM.sortBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      DOM.sortBtns.forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');
      state.currentSort = btn.dataset.sort;
      renderCryptoTable();
    });
  });

  // Retry button
  if (DOM.retryBtn) {
    DOM.retryBtn.addEventListener('click', fetchCoins);
  }

  // Voice controls
  DOM.speakBtn.addEventListener('click', speakBriefing);
  DOM.stopBtn.addEventListener('click', stopSpeaking);
  DOM.voiceRate.addEventListener('input', () => {
    DOM.voiceRateDisplay.textContent = `${parseFloat(DOM.voiceRate.value).toFixed(1)}x`;
  });

  // Newsletter
  DOM.newsletterForm.addEventListener('submit', handleNewsletterSubmit);

  // Fetch data
  await Promise.all([fetchCoins(), fetchGlobalData()]);
  startAutoRefresh();
}

// Boot
document.addEventListener('DOMContentLoaded', init);
