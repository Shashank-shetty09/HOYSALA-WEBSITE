let battingData=[],bowlingData=[];

function extractRuns(best){
    best = String(best).trim();
    let match = best.match(/^(\d+)/);
    return match ? parseInt(match[1]) : 0;
}

function extractWickets(spell){
    let parts = String(spell).split("-");
    return parseInt(parts[3]) || 0;
}

function extractRunsGiven(spell){
    let parts = String(spell).split("-");
    return parseInt(parts[2]) || 999;
}

fetch("HOYSALA_STATS.xlsx").then(r=>r.arrayBuffer()).then(d=>{
const wb=XLSX.read(d,{type:"array"});
battingData=XLSX.utils.sheet_to_json(wb.Sheets["BATTING"],{defval:""});
bowlingData=XLSX.utils.sheet_to_json(wb.Sheets["BOWLING"],{defval:""});
renderCurrentPage();
});




function getBattingBadge(p){
if(!p||p.MATCHES<15||p.STRIKERATE<110)
return "";
if(+p.STRIKERATE>=180)
return "🏏DESTROYER";
if(+p.STRIKERATE>=150)
return "🏏HARD HITTER";
if(+p.STRIKERATE>=130)
return "🏏ACCUMULATOR";
return "🏏CLASSICIST";
}

function getBowlingBadge(p){
if(!p||p.WICKETS<15)
return "";
if(+p.WICKETS>=100)
return "🥎SPEARHEAD";
if(+p.WICKETS>=80)
return "🥎VETERAN";
if(+p.WICKETS >= 60)
return "🥎STRIKE LEADER";
if(+p.ECONOMY<=8)
return "🥎ECONOMIST";
if(+p.WICKETS>=40)
return "🥎WILD CARD";
if(+p.WICKETS>=15)
return"🥎ASPIRANT";
}

function loadPlayers(){
let html="";

battingData.forEach(p=>{
let bowl=bowlingData.find(x=>x.BOWLER===p.BATSMAN);
let badge=getBattingBadge(p);
let bowlBadge=getBowlingBadge(bowl);

if(bowlBadge)
badge+=" "+bowlBadge;

let img=p.BATSMAN.toLowerCase()+".jpeg";

html+=`
<div class="player" data-player-name="${p.BATSMAN}">
<img src="${img}"
onerror="this.src='logo.png'">
<div class="playerInfo">
<h4>${p.BATSMAN}</h4>
<span class="badge">${badge}</span>
</div>
</div>
`;
});
pageContent.innerHTML=html;
}

function showPlayer(name){
let img=name.toLowerCase()+".jpeg";
let bat=battingData.find(x=>x.BATSMAN===name);
let bowl=bowlingData.find(x=>x.BOWLER===name);
let badge=getBattingBadge(bat);
let bowlBadge=getBowlingBadge(bowl);

if(bowlBadge)
badge += " " + bowlBadge;
pageContent.innerHTML=`
<button class="backBtn" onclick="window.location.href='players.html'">
← Back
</button>
<div class="profileTop">
<img src="${img}" onerror="this.src='logo.png'">
<div>
<div class="playerName">${name}</div>
<div class="badge">${badge}</div>
</div>
</div>
<div class="playerTabs">
<button class="active"
onclick="setActiveButton(this); showBatting('${name}')">BATTING</button>
<button onclick="setActiveButton(this); showBowling('${name}')">BOWLING</button>
</div>
<div id="playerData"></div>
`;
showBatting(name);
}

function setActiveButton(button){
    const parent = button.parentElement;
    parent.querySelectorAll("button").forEach(btn=>{
        btn.classList.remove("active");
    });
    button.classList.add("active");
}

function renderStats(obj,stats){
let h='<div class="statsGrid">';
stats.forEach(s=>h+=`<div class="card"><div class="label">${s[1]}</div><div class="value">${obj[s[0]]||0}</div></div>`);
h+='</div>';
document.getElementById("playerData").innerHTML=h;
}

function showBatting(name){
let p=battingData.find(x=>x.BATSMAN===name);
renderStats(p,[["MATCHES","Matches"],["INNINGS","Innings"],["RUNS","Runs"],["BALLS","Balls"],["NOT OUTS","Not Outs"],["STRIKERATE","SR"],["AVERAGE","Average"],["Balls/bdry","Balls/Boundary"],["FOURS","Fours"],["SIXES","Sixes"],["50","50s"],["BEST","Best"],["RECENT","Recent"]]);
}

function showBowling(name){
let p=bowlingData.find(x=>x.BOWLER===name);
if(!p){document.getElementById("playerData").innerHTML="<h3>No Bowling Data</h3>";return;}
renderStats(p,[["MATCHES","Matches"],["OVERS","Overs"],["WICKETS","Wickets"],["RUNS","Runs"],["ECONOMY","Economy"],["AVERAGE","Average"],["STRIKE RATE","SR"],["3-W","3W"],["5-W","5W"],["BEST","Best"],["RECENT","Recent"]]);
}


function loadLeaderboardPage(){
pageContent.innerHTML=`
<div class="subTabs">
<button class="active" onclick="setActiveButton(this); loadBattingBoard()">BATTING</button>
<button onclick="setActiveButton(this); loadBowlingBoard()">BOWLING</button>
<button onclick="setActiveButton(this); loadMVP()">MVP</button>
</div>
<div id="board"></div>`;
loadBattingBoard();
}

function loadBattingBoard(sortKey="RUNS"){
let s=[...battingData];
if(sortKey==="Balls/bdry"){
s.sort((a,b)=>
(+a[sortKey]||999)-(+b[sortKey]||999)
);
}else{
s.sort((a,b)=>
(+b[sortKey]||0)-(+a[sortKey]||0)
);
}
let h=`
<select onchange="loadBattingBoard(this.value)">
<option value="RUNS" ${sortKey==="RUNS"?"selected":""}>Runs</option>
<option value="STRIKERATE" ${sortKey==="STRIKERATE"?"selected":""}>Strike Rate</option>
<option value="AVERAGE" ${sortKey==="AVERAGE"?"selected":""}>Average</option>
<option value="FOURS" ${sortKey==="FOURS"?"selected":""}>Fours</option>
<option value="SIXES" ${sortKey==="SIXES"?"selected":""}>Sixes</option>
<option value="Balls/bdry" ${sortKey==="Balls/bdry"?"selected":""}>Balls/Boundary</option>
<option value="50" ${sortKey==="50"?"selected":""}>50s</option>
</select>
`;
let sortName={
RUNS:"Runs",
AVERAGE:"Average",
STRIKERATE:"Strike Rate",
FOURS:"Fours",
SIXES:"Sixes",
"50":"50s"
};
s.forEach((p,i)=>{
let img=p.BATSMAN.toLowerCase()+".jpeg";
h+=`
<div class="player" data-player-name="${p.BATSMAN}">
<img src="${img}"
onerror="this.src='logo.png'">
<div style="flex:1">
<b>#${i+1} ${p.BATSMAN}</b>
<br>
<small>Runs:${p.RUNS}</small>
<small>   SR:${p.STRIKERATE}</small>
<small>   AVG:${p.AVERAGE}</small>
</div>
<div>
<b>${p[sortKey]}</b>
</div>
</div>
`;
});
document.getElementById("board").innerHTML=h;
}

function loadBowlingBoard(sortKey="WICKETS"){
let s=[...bowlingData];
if(sortKey==="ECONOMY"||sortKey==="AVERAGE"||sortKey==="STRIKE RATE"){
s.sort((a,b)=>
(+a[sortKey]||999)-(+b[sortKey]||999)
);
}else{
s.sort((a,b)=>
(+b[sortKey]||0)-(+a[sortKey]||0)
);
}
let h=`
<select onchange="loadBowlingBoard(this.value)">
<option value="WICKETS" ${sortKey==="WICKETS"?"selected":""}>Wickets</option>
<option value="ECONOMY" ${sortKey==="ECONOMY"?"selected":""}>Economy</option>
<option value="AVERAGE" ${sortKey==="AVERAGE"?"selected":""}>Average</option>
<option value="STRIKE RATE" ${sortKey==="STRIKE RATE"?"selected":""}>Strike Rate</option>
<option value="3-W" ${sortKey==="3-W"?"selected":""}>3W</option>
<option value="5-W" ${sortKey==="5-W"?"selected":""}>5W</option>
</select>
`;
let sortName={
WICKETS:"Wickets",
ECONOMY:"Economy",
AVERAGE:"Average",
"STRIKE RATE":"Strike Rate",
"3-W":"3W",
"5-W":"5W"
};
s.forEach((p,i)=>{
let img=p.BOWLER.toLowerCase()+".jpeg";
h+=`
<div class="player" data-player-name="${p.BOWLER}">
<img src="${img}"
onerror="this.src='logo.png'">
<div style="flex:1">
<b>#${i+1} ${p.BOWLER}</b>
<br>
<small> Wickets:${p.WICKETS}</small>
<small> ECO:${p.ECONOMY}</small>
<small> SR:${p["STRIKE RATE"]}</small>
<small> AVG:${p.AVERAGE}</small>
</div>
<div>
<b>${p[sortKey]}</b>
</div>
</div>
`;
});
document.getElementById("board").innerHTML=h;
}

function loadMVP(){
    if(!battingData.length || !bowlingData.length){
    document.getElementById("board").innerHTML="<h3>Loading data...</h3>";
    return;
}
let players={};

battingData.forEach(p=>{
let battingScore =
(+p.RUNS || 0) * 1 +
(+p.AVERAGE || 0) * 2 +
(Math.max(0,(+p.STRIKERATE || 0) - 100)) * 2 +
(+p.FOURS || 0) * 1 +
(+p.SIXES || 0) * 2 +
(+p["50"] || 0) * 25;
if(!players[p.BATSMAN]){
players[p.BATSMAN]={
batting:0,
bowling:0
};
}
players[p.BATSMAN].batting=battingScore;
});

bowlingData.forEach(p=>{
let bowlingScore =
(+p.WICKETS || 0) * 20 +
(+p["3-W"] || 0) * 20 +
(+p["5-W"] || 0) * 50 +
Math.max(0,(12 - (+p.ECONOMY || 12))) * 20 +
Math.max(0,(15 - (+p["STRIKE RATE"] || 15))) * 10 +
Math.max(0,(20 - (+p.AVERAGE || 20))) * 10;
if(!players[p.BOWLER]){
players[p.BOWLER]={
batting:0,
bowling:0
};
}
players[p.BOWLER].bowling=bowlingScore;
});

let mvpList = Object.entries(players)
.map(([name,data]) => ({
name,
...data,
total:data.batting + data.bowling
}))
.sort((a,b)=>b.total-a.total);

mvpList.sort((a,b)=>b.total-a.total);

let s = mvpList;

let h="<h2>MVP Leaderboard</h2>";
s.forEach((p,i)=>{

let img=p.name.toLowerCase()+".jpeg";

let total=p.total;
h += `
<div class="player" data-player-name="${p.name}">
<img src="${img}" onerror="this.src='logo.png'">
<div style="flex:1">
<b>#${i+1} ${p.name}</b>
<br>
<small>
Batting: ${p.batting.toFixed(0)}
|
Bowling: ${p.bowling.toFixed(0)}
</small>
</div>
<div>
<b>${total.toFixed(0)}</b>
</div>
</div>
`;
});
document.getElementById("board").innerHTML=h;
}

function getNextMilestone(value, milestones){
    for(let m of milestones){
        if(value < m){
            return m;
        }
    }
    return null;
}

function getMilestoneScore(stat,target){

    if(stat==="Runs"){
        return {
            500:100,
            750:200,
            1000:300,
            1500:400,
            2000:500,
            2500:600
        }[target] || 0;
    }

    if(stat==="Wickets"){
        return {
            50:150,
            100:350,
            150:550,
            200:750,
            250:950
        }[target] || 0;
    }

    if(stat==="Sixes"){
        return {
            50:80,
            75:140,
            100:220,
            150:320,
            200:420,
            250:520
        }[target] || 0;
    }

    if(stat==="Fours"){
        return {
            50:60,
            75:120,
            100:180,
            150:260,
            200:340
        }[target] || 0;
    }

    if(stat==="Matches"){
        return {
            50:40,
            100:100,
            150:180,
            200:260,
            250:340
        }[target] || 0;
    }

    return 0;
}

function loadMilestones(){

let milestones = [];

/* RUNS */
battingData.forEach(p=>{
    let value = +p.RUNS || 0;
    let next = getNextMilestone(value,[500,750,1000,1500,2000,2500]);

    if(next && (next-value)<=50){
        milestones.push({
            player:p.BATSMAN,
            stat:"Runs",
            current:value,
            target:next,
            remaining:next-value,
            icon:"🏏"
        });
    }
});

/* WICKETS */
bowlingData.forEach(p=>{
    let value = +p.WICKETS || 0;
    let next = getNextMilestone(value,[50,100,150,200,250]);

    if(next && (next-value)<=12){
        milestones.push({
            player:p.BOWLER,
            stat:"Wickets",
            current:value,
            target:next,
            remaining:next-value,
            icon:"🥎"
        });
    }
});

/* SIXES */
battingData.forEach(p=>{
    let value = +p.SIXES || 0;
    let next = getNextMilestone(value,[50,100,150,200,250]);

    if(next && (next-value)<=10){
        milestones.push({
            player:p.BATSMAN,
            stat:"Sixes",
            current:value,
            target:next,
            remaining:next-value,
            icon:"💥"
        });
    }
});

/* FOURS */
battingData.forEach(p=>{
    let value = +p.FOURS || 0;
    let next = getNextMilestone(value,[50,100,150,200]);

    if(next && (next-value)<=10){
        milestones.push({
            player:p.BATSMAN,
            stat:"Fours",
            current:value,
            target:next,
            remaining:next-value,
            icon:"⚡"
        });
    }
});

/* MATCHES */
battingData.forEach(p=>{
    let value = +p.MATCHES || 0;
    let next = getNextMilestone(value,[50,100,150,200,250]);

    if(next && (next-value)<=6){
        milestones.push({
            player:p.BATSMAN,
            stat:"Matches",
            current:value,
            target:next,
            remaining:next-value,
            icon:"🎮"
        });
    }
});

milestones.sort((a,b)=>{

    let scoreA = getMilestoneScore(a.stat,a.target);
    let scoreB = getMilestoneScore(b.stat,b.target);

    if(scoreA !== scoreB)
        return scoreB - scoreA;

    return a.remaining - b.remaining;
});

let html = `
<h2>🎯 Milestone Tracker</h2>
<p>THESE GUYS ARE NEAR TO THESE MILESTONES</P>
<div class="statsGrid">
`;

milestones.forEach(m=>{
    html += `
    <div class="card">
        <div class="label">
            ${m.icon} ${m.stat}
        </div>

        <div class="value">
            ${m.current}/${m.target}
        </div>

        <p>
            <b>${m.player}</b>
        </p>

        <p>
            ${m.remaining} Away
        </p>
    </div>
    `;
});

html += "</div>";

pageContent.innerHTML = html;
}

function loadDetails(){

pageContent.innerHTML=`

<h2>🏅 Badge Details</h2>

<div class="card">
<h3>Batting Badges</h3>

<p><b>🏏 DESTROYER</b><br>
Strike Rate >180 and Matches ≥ 15</p>

<p><b>🏏 HARD HITTER</b><br>
Strike Rate >150 and Matches ≥ 15</p>

<p><b>🏏 ACCUMULATOR</b><br>
Strike Rate >130 and Matches ≥ 15</p>

<p><b>🏏 CLASSICIST</b><br>
Strike Rate >110 and Matches ≥ 15</p>
</div>

<br>

<div class="card">
<h3>Bowling Badges</h3>

<p><b>🥎 SPEARHEAD</b><br>
100 or more wickets</p>

<p><b>🥎 VETERAN</b><br>
80 or more wickets</p>

<p><b>🥎 STRIKE LEADER</b><br>
60 or more wickets</p>

<p><b>🥎 ECONOMIST</b><br>
Economy ≤ 8</p>

<p><b>🥎 WILD CARD</b><br>
40 or more wickets</p>

<p><b>🥎 ASPIRANT</b><br>
15 or more wickets</p>

</div>

`;
}

function oversToBalls(overs){

    overs = String(overs);

    let parts = overs.split(".");

    let completeOvers = parseInt(parts[0]) || 0;
    let balls = parseInt(parts[1]) || 0;

    return completeOvers * 6 + balls;
}

function loadRecords(){
if(!battingData.length || !bowlingData.length){
    pageContent.innerHTML="<h3>Loading data...</h3>";
    return;
}

let topRunScorer =
    battingData.reduce((a,b)=>
        +a.RUNS > +b.RUNS ? a:b);

let topWicketTaker =
    bowlingData.reduce((a,b)=>
        +a.WICKETS > +b.WICKETS ? a:b);

let bestSR =
    battingData.reduce((a,b)=>
        +a.STRIKERATE > +b.STRIKERATE ? a:b);

let bestEconomy =
    bowlingData.reduce((a,b)=>
        +a.ECONOMY < +b.ECONOMY ? a:b);

let mostFifties =
    battingData.reduce((a,b)=>
        +a["50"] > +b["50"] ? a:b);

let mostFifers =
    bowlingData.reduce((a,b)=>
        +a["5-W"] > +b["5-W"] ? a : b
    );

let mostThreefers =
    bowlingData.reduce((a,b)=>
        +a["3-W"] > +b["3-W"] ? a : b
    );

let bestAverage =
    battingData.reduce((a,b)=>
        +a.AVERAGE > +b.AVERAGE ? a:b);

let bestInnings = battingData.reduce((a,b)=>{
    return extractRuns(a.BEST) > extractRuns(b.BEST) ? a : b;
});

let bestSpell = bowlingData.reduce((a,b)=>{
    let wa = extractWickets(a.BEST);
    let wb = extractWickets(b.BEST);

    if(wa !== wb)
        return wa > wb ? a : b;

    return extractRunsGiven(a.BEST) < extractRunsGiven(b.BEST) ? a : b;
});

let totalInnings = 0;
let totalRuns = 0;
let totalBalls = 0;
let totalNotOuts = 0;
let totalFours = 0;
let totalSixes = 0;
let totalFifties = 0;

battingData.forEach(p=>{
    totalInnings += +p.INNINGS || 0;
    totalRuns += +p.RUNS || 0;
    totalBalls += +p.BALLS || 0;
    totalNotOuts += +p["NOT OUTS"] || 0;
    totalFours += +p.FOURS || 0;
    totalSixes += +p.SIXES || 0;
    totalFifties += +p["50"] || 0;
});

let battingAverage =
totalRuns / Math.max(1,(totalInnings-totalNotOuts));

let battingStrikeRate =
(totalRuns * 100) / Math.max(1,totalBalls);

let ballsPerBoundary =
totalBalls / Math.max(1,(totalFours + totalSixes));

/* BOWLING TOTALS */

let totalBallsBowled = 0;
let totalBowlingRuns = 0;
let totalWickets = 0;
let total3W = 0;
let total5W = 0;

bowlingData.forEach(p=>{
    totalBallsBowled += oversToBalls(p.OVERS);
    totalBowlingRuns += +p.RUNS || 0;
    totalWickets += +p.WICKETS || 0;
    total3W += +p["3-W"] || 0;
    total5W += +p["5-W"] || 0;
});

let bowlingEconomy =
(totalBowlingRuns * 6) /
Math.max(1,totalBallsBowled);

let bowlingAverage =
totalBowlingRuns /
Math.max(1,totalWickets);

let bowlingStrikeRate =
totalBallsBowled /
Math.max(1,totalWickets);

let totalOversDisplay =
Math.floor(totalBallsBowled / 6) +
"." +
(totalBallsBowled % 6);

pageContent.innerHTML=`

<h2>📊 Team Records</h2>

<div class="statsGrid">

<div class="card">
<div class="label">Highest Runs</div>
<div class="value">${topRunScorer.RUNS}</div>
<p>${topRunScorer.BATSMAN}</p>
</div>

<div class="card">
<div class="label">Highest Wickets</div>
<div class="value">${topWicketTaker.WICKETS}</div>
<p>${topWicketTaker.BOWLER}</p>
</div>

<div class="card">
<div class="label">Best Strike Rate</div>
<div class="value">${bestSR.STRIKERATE}</div>
<p>${bestSR.BATSMAN}</p>
</div>

<div class="card">
<div class="label">Best Economy</div>
<div class="value">${bestEconomy.ECONOMY}</div>
<p>${bestEconomy.BOWLER}</p>
</div>

<div class="card">
<div class="label">Most 50s</div>
<div class="value">${mostFifties["50"]}</div>
<p>${mostFifties.BATSMAN}</p>
</div>

<div class="card">
<div class="label">Most 5-Wickets</div>
<div class="value">${mostFifers["5-W"]}</div>
<p>${mostFifers.BOWLER}</p>
</div>

<div class="card">
<div class="label">Most 3-Wickets</div>
<div class="value">${mostThreefers["3-W"]}</div>
<p>${mostThreefers.BOWLER}</p>
</div>

<div class="card">
<div class="label">Best Batting Average</div>
<div class="value">${bestAverage.AVERAGE}</div>
<p>${bestAverage.BATSMAN}</p>
</div>

<div class="card">
<div class="label">Best Innings</div>
<div class="value">${bestInnings.BEST}</div>
<p>${bestInnings.BATSMAN}</p>
</div>

<div class="card">
<div class="label">Best Spell</div>
<div class="value">${bestSpell.BEST}</div>
<p>${bestSpell.BOWLER}</p>
</div>

<div class="card">
<div class="label">First 1000runs</div>
<p>HONNUR SWAMY</p>
</div>

<div class="card">
<div class="label">First 100 Wickets</div>
<p>SHASHANK SHETTY</p>
</div>

<div class="card">
<div class="label">First 100 6S</div>
<p>HONNUR SWAMY</p>
</div>

<div class="card">
<div class="label">First Century</div>
<p>HONNUR SWAMY</p>
</div>

<div class="card">
<div class="label">First 100 Fours</div>
<p>HONNUR SWAMY</p>
</div>

<div class="card">
<div class="label">First 5FER</div>
<p>SHASHANK SHETTY</p>
</div>

<div class="card">
<div class="label">First HATTRICK</div>
<p>SHASHANK SHETTY</p>
</div>

</div>

<div class="box">
    <p>MATCHES:CALCULATING</P>
    <p>WON:CALCULATING</p>
    <p>LOST:CALCULATING</p>
    </div>
<h2>🏏 TEAM BATTING STATS</h2>

<div class="statsGrid">

<div class="card">
<div class="label">Total Runs</div>
<div class="value">${totalRuns}</div>
</div>

<div class="card">
<div class="label">Batting Average</div>
<div class="value">${battingAverage.toFixed(2)}</div>
</div>

<div class="card">
<div class="label">Batting Strike Rate</div>
<div class="value">${battingStrikeRate.toFixed(2)}</div>
</div>

<div class="card">
<div class="label">Balls / Boundary</div>
<div class="value">${ballsPerBoundary.toFixed(2)}</div>
</div>

<div class="card">
<div class="label">Fours</div>
<div class="value">${totalFours}</div>
</div>

<div class="card">
<div class="label">Sixes</div>
<div class="value">${totalSixes}</div>
</div>

<div class="card">
<div class="label">50s</div>
<div class="value">${totalFifties}</div>
</div>
</div>
<h2 style="margin-top:30px;">🥎 TEAM BOWLING STATS</h2>
<div class="statsGrid">
<div class="card">
<div class="label">Overs Bowled</div>
<div class="value">${totalOversDisplay}</div>
</div>

<div class="card">
<div class="label">Wickets</div>
<div class="value">${totalWickets}</div>
</div>

<div class="card">
<div class="label">Bowling Economy</div>
<div class="value">${bowlingEconomy.toFixed(2)}</div>
</div>

<div class="card">
<div class="label">Bowling Average</div>
<div class="value">${bowlingAverage.toFixed(2)}</div>
</div>

<div class="card">
<div class="label">Bowling Strike Rate</div>
<div class="value">${bowlingStrikeRate.toFixed(2)}</div>
</div>

<div class="card">
<div class="label">3-Wicket Hauls</div>
<div class="value">${total3W}</div>
</div>

<div class="card">
<div class="label">5-Wicket Hauls</div>
<div class="value">${total5W}</div>
</div>
</div>
</div>
`; 
}

function openImage(src){
    document.getElementById("imageModal").style.display="flex";
    document.getElementById("modalImg").src=src;
}

function closeImage(){
    document.getElementById("imageModal").style.display="none";
}

function loadGallery(){

let html = `
<h2>📸 Hoysala Gallery</h2>
<div class="galleryGrid">
`;

for(let i=1;i<=250;i++){

    html += `
    <div class="galleryCard">
        <img src="gallery${i}.jpg"
             onerror="this.parentElement.style.display='none'"
             onclick="openImage('gallery${i}.jpg')">
    </div>
    `;
}

html += "</div>";

pageContent.innerHTML = html;
}

function setActivePage(){
    const page = document.body.dataset.page;
    document.querySelectorAll(".mainTabs a").forEach(a=>{
        a.classList.toggle("active", a.dataset.page===page);
    });
}

function renderCurrentPage(){
    setActivePage();

    const page = document.body.dataset.page;
    if(page==="players") loadPlayers();
    if(page==="leaderboard") loadLeaderboardPage();
    if(page==="records") loadRecords();
    if(page==="milestones") loadMilestones();
    if(page==="gallery") loadGallery();
    if(page==="details") loadDetails();

    // Open a player profile directly from players.html?name=...
    if(page==="players"){
        const name = new URLSearchParams(location.search).get("name");
        if(name) showPlayer(decodeURIComponent(name));
    }
}

function navigateToPlayer(name){
    window.location.href = "players.html?name=" + encodeURIComponent(name);
}

document.addEventListener("DOMContentLoaded",()=>{
    window.pageContent = document.getElementById("content");

    // Preserve normal browser navigation. No SPA history manipulation is used.
    // Clicking a player creates a real URL.
    document.addEventListener("click", e=>{
        const player = e.target.closest(".player[data-player-name]");
        if(player){
            e.preventDefault();
            navigateToPlayer(player.dataset.playerName);
        }
    });

    renderCurrentPage();
});
