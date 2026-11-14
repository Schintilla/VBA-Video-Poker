const SUITS = [
  { name: 'Spades', symbol: '♠', color: 'black' },
  { name: 'Hearts', symbol: '♥', color: 'red' },
  { name: 'Diamonds', symbol: '♦', color: 'red' },
  { name: 'Clubs', symbol: '♣', color: 'black' },
];
const RANKS = [null, 'A', '2', '3', '4', '5', '6', '7', '8', '9', '10', 'J', 'Q', 'K'];
const PAYOUTS = [
  ['Royal Flush', 250], ['Straight Flush', 50], ['Four of a Kind', 25],
  ['Full House', 9], ['Flush', 6], ['Straight', 4], ['Three of a Kind', 3],
  ['Two Pair', 2], ['Jacks or Better', 1],
];
const state = {
  credits: 100, bet: 5, hand: [], deck: [], held: new Set(), phase: 'ready',
  plays: 0, wins: 0, creditsWon: 0, results: Object.fromEntries(PAYOUTS.map(([name]) => [name, 0])),
  bankroll: [100], auto: false,
};

const $ = (selector) => document.querySelector(selector);
const cardsEl = $('#cards');

function newDeck() {
  const deck = [];
  for (const suit of SUITS) for (let rank = 1; rank <= 13; rank++) deck.push({ rank, suit });
  for (let i = deck.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [deck[i], deck[j]] = [deck[j], deck[i]];
  }
  return deck;
}

function deal() {
  if (state.phase !== 'ready') return false;
  state.bet = Number($('#bet').value);
  if (state.credits < state.bet) {
    showMessage('Not enough credits for that bet. Lower your bet or reset the session.');
    return false;
  }
  state.credits -= state.bet;
  state.deck = newDeck();
  state.hand = state.deck.splice(0, 5);
  state.held.clear();
  state.phase = 'draw';
  render();
  showMessage('Select cards to hold, then draw.');
  return true;
}

function rankValue(rank) { return rank === 1 ? 14 : rank; }

function evaluate(hand) {
  const ranks = hand.map(card => rankValue(card.rank)).sort((a, b) => a - b);
  const counts = new Map();
  for (const rank of ranks) counts.set(rank, (counts.get(rank) || 0) + 1);
  const groups = [...counts.values()].sort((a, b) => b - a);
  const flush = hand.every(card => card.suit === hand[0].suit);
  const unique = [...new Set(ranks)];
  const straight = unique.length === 5 && (unique[4] - unique[0] === 4 || unique.join(',') === '2,3,4,5,14');
  const royal = straight && ranks[0] === 10 && ranks[4] === 14;
  if (flush && royal) return 'Royal Flush';
  if (flush && straight) return 'Straight Flush';
  if (groups[0] === 4) return 'Four of a Kind';
  if (groups[0] === 3 && groups[1] === 2) return 'Full House';
  if (flush) return 'Flush';
  if (straight) return 'Straight';
  if (groups[0] === 3) return 'Three of a Kind';
  if (groups[0] === 2 && groups[1] === 2) return 'Two Pair';
  if (groups[0] === 2) {
    const pairRank = [...counts].find(([, count]) => count === 2)[0];
    if (pairRank >= 11 || pairRank === 14) return 'Jacks or Better';
  }
  return 'No win';
}

function settle() {
  if (state.phase !== 'draw') return;
  const held = state.hand.filter((_, index) => state.held.has(index));
  const replacements = state.deck.splice(0, 5 - held.length);
  state.hand = state.hand.filter((_, index) => state.held.has(index)).concat(replacements);
  const result = evaluate(state.hand);
  let payout = 0;
  if (result !== 'No win') {
    payout = result === 'Royal Flush' && state.bet === 5 ? 4000 : PAYOUTS.find(([name]) => name === result)[1] * state.bet;
    state.credits += payout;
    state.wins++;
    state.creditsWon += payout;
    state.results[result]++;
  }
  state.plays++;
  state.bankroll.push(state.credits);
  state.phase = 'ready';
  state.held.clear();
  render();
  showMessage(result === 'No win' ? 'No winning hand this time.' : `${result}! You won ${payout} ${payout === 1 ? 'credit' : 'credits'}.`);
}

function toggleHold(index) {
  if (state.phase !== 'draw') return;
  if (state.held.has(index)) state.held.delete(index);
  else state.held.add(index);
  renderCards();
}

function chooseHolds(hand) {
  const ranks = hand.map(card => rankValue(card.rank));
  const counts = new Map();
  ranks.forEach(rank => counts.set(rank, (counts.get(rank) || 0) + 1));
  const made = evaluate(hand);
  if (made !== 'No win') return [0, 1, 2, 3, 4];

  // Keep four cards to a royal flush, then four to a flush or open-ended straight.
  for (const suit of SUITS) {
    const suited = hand.map((card, i) => card.suit === suit ? i : -1).filter(i => i >= 0);
    const royalCards = suited.filter(i => [10, 11, 12, 13, 14].includes(ranks[i]));
    if (royalCards.length === 4) return royalCards;
  }
  const suitGroups = SUITS.map(suit => hand.map((card, i) => card.suit === suit ? i : -1).filter(i => i >= 0));
  const fourFlush = suitGroups.find(group => group.length === 4);
  if (fourFlush) return fourFlush;
  for (let start = 2; start <= 10; start++) {
    const run = [start, start + 1, start + 2, start + 3];
    const indices = run.map(rank => ranks.indexOf(rank));
    if (indices.every(index => index >= 0)) return indices;
  }
  const pairs = [...counts].filter(([, count]) => count === 2).map(([rank]) => rank);
  if (pairs.length) {
    const keep = pairs.filter(rank => rank >= 11 || rank === 14);
    if (keep.length) return hand.map((card, i) => keep.includes(ranks[i]) ? i : -1).filter(i => i >= 0);
    if (pairs.length === 2) return hand.map((card, i) => pairs.includes(ranks[i]) ? i : -1).filter(i => i >= 0);
    return hand.map((card, i) => ranks[i] === pairs[0] ? i : -1).filter(i => i >= 0);
  }
  const trips = [...counts].filter(([, count]) => count === 3).map(([rank]) => rank);
  if (trips.length) return hand.map((card, i) => ranks[i] === trips[0] ? i : -1).filter(i => i >= 0);
  const high = hand.map((card, i) => ({ i, rank: ranks[i] })).filter(card => card.rank >= 11 || card.rank === 14).sort((a, b) => b.rank - a.rank)[0];
  return high ? [high.i] : [];
}

function runAutoPlay() {
  if (state.auto) {
    state.auto = false;
    $('#auto-button').textContent = 'Run';
    showMessage('Auto-play stopped.');
    return;
  }
  const requested = Math.max(1, Math.min(10000, Number($('#auto-count').value) || 1));
  let remaining = requested;
  state.auto = true;
  $('#auto-button').textContent = 'Stop';
  const tick = () => {
    if (!state.auto) return;
    let batch = 0;
    while (remaining > 0 && batch < 35 && state.credits >= Number($('#bet').value)) {
      if (!deal()) break;
      state.held = new Set(chooseHolds(state.hand));
      settle();
      remaining--;
      batch++;
    }
    if (remaining <= 0 || state.credits < Number($('#bet').value)) {
      state.auto = false;
      $('#auto-button').textContent = 'Run';
      showMessage(remaining <= 0 ? `Auto-play complete: ${requested} hands.` : 'Auto-play stopped: not enough credits for another hand.');
      return;
    }
    setTimeout(tick, 0);
  };
  tick();
}

function cardMarkup(card, index) {
  const held = state.held.has(index);
  const color = card.suit.color;
  const label = `${RANKS[card.rank]} of ${card.suit.name}${held ? ', held' : ''}`;
  return `<button class="playing-card ${color} ${held ? 'held' : ''}" type="button" data-index="${index}" aria-label="${label}" aria-pressed="${held}" ${state.phase !== 'draw' ? 'disabled' : ''}>
    <span class="corner top"><b>${RANKS[card.rank]}</b><span>${card.suit.symbol}</span></span>
    <span class="card-suit" aria-hidden="true">${card.suit.symbol}</span>
    <span class="corner bottom"><b>${RANKS[card.rank]}</b><span>${card.suit.symbol}</span></span>
    ${held ? '<span class="hold-tag">HELD</span>' : ''}
  </button>`;
}

function renderCards() {
  if (!state.hand.length) {
    cardsEl.innerHTML = '<div class="card-placeholder">Your five cards will appear here</div>';
  } else {
    cardsEl.innerHTML = state.hand.map(cardMarkup).join('');
    cardsEl.querySelectorAll('.playing-card').forEach(button => button.addEventListener('click', () => toggleHold(Number(button.dataset.index))));
  }
}

function renderStats() {
  $('#credits').textContent = state.credits.toLocaleString();
  $('#plays').textContent = state.plays.toLocaleString();
  $('#wins').textContent = state.wins.toLocaleString();
  $('#win-rate').textContent = `${state.plays ? (state.wins * 100 / state.plays).toFixed(1) : '0'}%`;
  $('#credits-won').textContent = state.creditsWon.toLocaleString();
  $('#deal-button').textContent = state.phase === 'draw' ? 'Draw' : 'Deal';
  $('#deal-button').disabled = state.phase === 'ready' && state.credits < Number($('#bet').value);
  $('#hold-button').disabled = state.phase !== 'draw';
  $('#hold-button').textContent = state.phase === 'draw' ? `Hold winning hand (${chooseHolds(state.hand).length})` : 'Hold winning hand';
  $('#hand-status').textContent = state.phase === 'draw' ? 'Choose cards to hold' : 'Ready when you are';
  drawCharts();
}

function render() { renderCards(); renderStats(); }
function showMessage(message) { $('#message').textContent = message; }

function drawCharts() {
  drawLineChart($('#bankroll-chart'), state.bankroll);
  drawBarChart($('#hands-chart'), state.results);
}

function setupCanvas(canvas) {
  const rect = canvas.getBoundingClientRect();
  const ratio = window.devicePixelRatio || 1;
  canvas.width = Math.max(1, Math.floor(rect.width * ratio));
  canvas.height = Math.max(1, Math.floor(rect.height * ratio));
  const context = canvas.getContext('2d');
  context.scale(ratio, ratio);
  return { context, width: rect.width, height: rect.height };
}

function drawLineChart(canvas, values) {
  if (!canvas.getBoundingClientRect().width) return;
  const { context: ctx, width, height } = setupCanvas(canvas);
  ctx.clearRect(0, 0, width, height);
  const pad = { left: 36, right: 12, top: 14, bottom: 24 };
  const min = Math.min(...values, 0), max = Math.max(...values, 100), span = Math.max(1, max - min);
  ctx.strokeStyle = '#e6ebe7'; ctx.lineWidth = 1;
  for (let i = 0; i < 4; i++) {
    const y = pad.top + (height - pad.top - pad.bottom) * i / 3;
    ctx.beginPath(); ctx.moveTo(pad.left, y); ctx.lineTo(width - pad.right, y); ctx.stroke();
  }
  ctx.fillStyle = '#78877e'; ctx.font = '11px system-ui'; ctx.textAlign = 'right';
  ctx.fillText(Math.round(max).toString(), pad.left - 7, pad.top + 4);
  ctx.fillText(Math.round(min).toString(), pad.left - 7, height - pad.bottom + 4);
  const plotW = width - pad.left - pad.right, plotH = height - pad.top - pad.bottom;
  ctx.beginPath();
  values.forEach((value, i) => {
    const x = pad.left + (values.length === 1 ? plotW : plotW * i / (values.length - 1));
    const y = pad.top + plotH * (max - value) / span;
    if (i === 0) ctx.moveTo(x, y); else ctx.lineTo(x, y);
  });
  ctx.strokeStyle = '#15966b'; ctx.lineWidth = 2.5; ctx.stroke();
}

function drawBarChart(canvas, values) {
  if (!canvas.getBoundingClientRect().width) return;
  const { context: ctx, width, height } = setupCanvas(canvas);
  ctx.clearRect(0, 0, width, height);
  const entries = Object.entries(values);
  const max = Math.max(1, ...entries.map(([, value]) => value));
  const pad = { left: 10, right: 8, top: 12, bottom: 46 };
  const plotH = height - pad.top - pad.bottom;
  const gap = 7, barW = (width - pad.left - pad.right - gap * (entries.length - 1)) / entries.length;
  entries.forEach(([label, value], i) => {
    const x = pad.left + i * (barW + gap), barH = plotH * value / max;
    ctx.fillStyle = value ? '#15966b' : '#dfe9e3';
    ctx.beginPath(); ctx.roundRect(x, pad.top + plotH - barH, Math.max(2, barW), Math.max(2, barH), 3); ctx.fill();
    ctx.fillStyle = '#78877e'; ctx.font = '10px system-ui'; ctx.textAlign = 'center';
    ctx.save(); ctx.translate(x + barW / 2, height - 4); ctx.rotate(-Math.PI / 4); ctx.fillText(label.replace(' of a Kind', 'K').replace('Straight Flush', 'Str. Flush').replace('Royal Flush', 'Royal').replace('Jacks or Better', 'J or B'), 0, 0); ctx.restore();
  });
}

function resetStats() {
  state.credits = 100; state.hand = []; state.deck = []; state.held.clear(); state.phase = 'ready';
  state.plays = 0; state.wins = 0; state.creditsWon = 0;
  state.results = Object.fromEntries(PAYOUTS.map(([name]) => [name, 0]));
  state.bankroll = [100];
  state.auto = false; $('#auto-button').textContent = 'Run';
  render(); showMessage('Session reset. Choose your bet, then deal.');
}

$('#deal-button').addEventListener('click', () => state.phase === 'ready' ? deal() : settle());
$('#hold-button').addEventListener('click', () => { if (state.phase === 'draw') { state.held = new Set(chooseHolds(state.hand)); renderCards(); } });
$('#auto-button').addEventListener('click', runAutoPlay);
$('#reset-button').addEventListener('click', resetStats);
$('#bet').addEventListener('change', renderStats);
$('#paytable-list').innerHTML = PAYOUTS.map(([name, value]) => `<div class="pay-row"><span>${name}</span><strong>${name === 'Royal Flush' ? '250× / 4,000×*' : `${value}×`}</strong></div>`).join('');
window.addEventListener('resize', drawCharts);
render();
