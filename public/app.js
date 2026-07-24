const zone = "Asia/Hong_Kong";
const isoFormatter = new Intl.DateTimeFormat("en-CA", { timeZone: zone, year: "numeric", month: "2-digit", day: "2-digit" });
const longFormatter = new Intl.DateTimeFormat("en-GB", { timeZone: zone, weekday: "long", day: "numeric", month: "long", year: "numeric" });
const shortFormatter = new Intl.DateTimeFormat("en-GB", { timeZone: zone, weekday: "short", day: "numeric", month: "short" });
const isoToday = () => { const parts = Object.fromEntries(isoFormatter.formatToParts(new Date()).filter(({ type }) => type !== "literal").map(({ type, value }) => [type, value])); return `${parts.year}-${parts.month}-${parts.day}`; };
const addDays = (date, days) => new Date(`${date}T00:00:00Z`).getTime() + days * 86400000;
const dateKey = (time) => new Date(time).toISOString().slice(0, 10);
const readable = (key, formatter = longFormatter) => formatter.format(new Date(`${key}T12:00:00Z`));
const detail = document.querySelector("#detail");
const grid = document.querySelector("#date-grid");
function mealInfo(entry) { if (!entry) return { students: null, meals: null, method: "No meal plan has been entered for this date." }; if (entry.expectedMeals !== undefined) return { students: entry.students, meals: entry.expectedMeals, method: "Recorded forecast" }; return { students: entry.students, meals: Math.ceil(entry.students * 0.65), method: entry.students === 0 ? "No school lunch planned" : "Estimated at 65% of planned students" }; }
function showDetail(key, entry) { const info = mealInfo(entry); detail.innerHTML = `<p class="eyebrow">SELECTED DATE</p><h2>${readable(key)}</h2><div class="stats"><div><span>Expected meals</span><strong>${info.meals === null ? "—" : info.meals.toLocaleString()}</strong></div><div><span>Students planned</span><strong>${info.students === null ? "—" : info.students.toLocaleString()}</strong></div></div><p class="method">${info.method}</p>`; }
async function init() {
  const [data, status] = await Promise.all([fetch("./meal-data.json").then((r) => r.json()), fetch("./daily-status.json").then((r) => r.json())]);
  const start = isoToday(); const keys = Array.from({ length: 8 }, (_, i) => dateKey(addDays(start, i)));
  document.querySelector("#range-label").textContent = `${readable(keys[0], shortFormatter)} – ${readable(keys.at(-1), shortFormatter)} · Hong Kong time`;
  document.querySelector("#updated-label").textContent = `Refreshed ${new Intl.DateTimeFormat("en-GB", { dateStyle: "medium", timeStyle: "short", timeZone: zone }).format(new Date(status.updatedAt))}`;
  keys.forEach((key, index) => { const info = mealInfo(data.mealPlan[key]); const button = document.createElement("button"); button.type = "button"; button.className = `date-button ${index === 0 ? "selected" : ""}`; button.innerHTML = `<span>${index === 0 ? "Today" : readable(key, shortFormatter)}</span><strong>${info.meals === null ? "—" : info.meals.toLocaleString()} meals</strong>`; button.addEventListener("click", () => { document.querySelectorAll(".date-button").forEach((item) => item.classList.remove("selected")); button.classList.add("selected"); showDetail(key, data.mealPlan[key]); }); grid.append(button); });
  showDetail(keys[0], data.mealPlan[keys[0]]);
}
init().catch(() => { detail.innerHTML = "<h2>Could not load the meal plan.</h2><p>Please refresh the page.</p>"; });
