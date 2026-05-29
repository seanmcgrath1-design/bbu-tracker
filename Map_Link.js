const MY_MAP_SPREADSHEET_ID = "1Ada_FMW6YmE25puTyjYA4QYWUxmlFLpTptEdgviUif0";
const MY_MAP_SHEET_NAME = "Daily Data Dump";
const FILTER_SHEET_NAME = "Site Detail"; 
const HUB_SHEET_NAME = "Hub Coordinates"; 

function findCol(rawArray, keyword1, keyword2 = "") {
  for (let r = 0; r < rawArray.length; r++) {
    let row = rawArray[r].map(h => String(h).toLowerCase().trim());
    let idx = row.findIndex(h => h.includes(keyword1) && (keyword2 === "" || h.includes(keyword2)));
    if (idx > -1) return { col: idx, row: r };
  }
  return { col: -1, row: -1 };
}

function fuzzyHubMatch(name) {
  let n = String(name).toUpperCase();
  n = n.replace(/\bAND\b/g, "").replace(/\bSTREET\b/g, "ST").replace(/\bROAD\b/g, "RD").replace(/\bAVENUE\b/g, "AVE");
  n = n.replace(/\bSOUTH\b/g, "S").replace(/\bNORTH\b/g, "N").replace(/\bEAST\b/g, "E").replace(/\bWEST\b/g, "W"); 
  n = n.replace(/[^A-Z0-9]/g, "");   
  if (n.startsWith("CHICAGO")) n = n.substring(7); else if (n.startsWith("CH")) n = n.substring(2); 
  return n;
}

function getMapData() {
  try {
    const ss = SpreadsheetApp.openById(MY_MAP_SPREADSHEET_ID);
    const dataSheet = ss.getSheetByName(MY_MAP_SHEET_NAME);
    const filterSheet = ss.getSheetByName(FILTER_SHEET_NAME);
    const hubSheet = ss.getSheetByName(HUB_SHEET_NAME);
    
    let fuzzyHubCoords = {};
    if (hubSheet) {
      const hubData = hubSheet.getDataRange().getDisplayValues();
      const hubRawTop = hubData.slice(0, 10);
      let hNameData = findCol(hubRawTop, "location");
      let hNameCol = hNameData.col; 
      let hLatCol = findCol(hubRawTop, "lat").col;
      let hLngCol = findCol(hubRawTop, "long").col; 
      if (hLngCol === -1) hLngCol = findCol(hubRawTop, "lon").col;

      if (hNameCol > -1 && hLatCol > -1 && hLngCol > -1) {
        for (let i = hNameData.row + 1; i < hubData.length; i++) {
          let hLat = parseFloat(hubData[i][hLatCol]); 
          let hLng = parseFloat(hubData[i][hLngCol]);
          let rawName = String(hubData[i][hNameCol]);
          
          // STRICTLY LOOK AT COLUMN E (Index 4) FOR THE HUB TYPE
          let hubTypeStr = (hubData[i].length > 4) ? String(hubData[i][4]).toUpperCase() : "";

          if (rawName !== "" && !isNaN(hLat) && !isNaN(hLng)) {
            // Check if it's an iEN hub (make sure it doesn't accidentally catch "NON IEN")
            let isIen = hubTypeStr.includes("IEN") && !hubTypeStr.includes("NON");
            fuzzyHubCoords[fuzzyHubMatch(rawName)] = { lat: hLat, lng: hLng, isIen: isIen };
          }
        }
      }
    }

    const filterLastRow = filterSheet.getLastRow();
    const filterLastCol = filterSheet.getLastColumn();
    const filterData = filterSheet.getRange(1, 1, filterLastRow, filterLastCol).getDisplayValues();
    const filterFormulas = filterSheet.getRange(1, 1, filterLastRow, filterLastCol).getFormulas();
    const filterBackgrounds = filterSheet.getRange(1, 1, filterLastRow, filterLastCol).getBackgrounds();
    
    const filterRawTop = filterData.slice(0, 10);
    let mIdData = findCol(filterRawTop, "fuze", "id");
    let filterHeaderRow = mIdData.row; let filterMatchIdx = mIdData.col;
    let filterBbuCol = findCol(filterRawTop, "bbu").col; let filterHubCol = findCol(filterRawTop, "hub").col;
    let filterOutCol = findCol(filterRawTop, "pairing").col; let filterSectorIdx = findCol(filterRawTop, "sector").col;

    let validSites = new Map();
    let frontendHubCoords = {}; 
    
    const dumpRawTop = dataSheet.getRange(1, 1, 10, dataSheet.getLastColumn()).getValues();
    let actIdx = -1, intIdx = -1, rdyIdx = -1;
    
    for (let r = 0; r < dumpRawTop.length; r++) {
      for (let c = 0; c < dumpRawTop[r].length; c++) {
        let clean = String(dumpRawTop[r][c]).toLowerCase().replace(/[^a-z0-9]/g, '');
        if (clean === "inserviceactivationa") actIdx = c;
        else if (clean === "bbintegrationcompleteda") intIdx = c;
        else if (clean === "bbintegrationreadya") rdyIdx = c;
      }
    }

    for (let r = filterHeaderRow + 1; r < filterData.length; r++) {
      let id = String(filterData[r][filterMatchIdx]).toLowerCase().replace(/\D/g, "");
      if (id && id !== "") {
        let rawBbu = (filterBbuCol !== -1) ? String(filterData[r][filterBbuCol]) : "";
        let bbuClean = rawBbu.toUpperCase().replace(/[^A-Z0-9]/g, ""); 
        let rawHub = (filterHubCol !== -1) ? String(filterData[r][filterHubCol]) : "OTHER";
        let displayHub = rawHub.replace(/\s+/g, " ").trim().toUpperCase();
        if (displayHub === "") displayHub = "UNASSIGNED HUB";
        
        let crushedHub = fuzzyHubMatch(rawHub);
        if (fuzzyHubCoords[crushedHub]) frontendHubCoords[displayHub] = fuzzyHubCoords[crushedHub];
        
        let color = (filterSectorIdx !== -1) ? filterBackgrounds[r][filterSectorIdx] : '#3388ff';
        if (color === '#ffffff' || !color) color = filterBackgrounds[r][filterOutCol]; 
        if (color === '#ffffff' || !color) color = '#3388ff';
        
        let clusterId = (bbuClean !== "") ? "BBU_" + bbuClean : (String(filterFormulas[r][filterOutCol]).match(/HYPERLINK\("([^"]+)"/i) ? "URL_" + String(filterFormulas[r][filterOutCol]).match(/HYPERLINK\("([^"]+)"/i)[1].replace(/[^A-Za-z0-9]/g, "").substring(0, 30) : "STANDALONE_" + r);

        validSites.set(id, { color: color, clusterId: clusterId, hub: displayHub, cleanBbu: rawBbu });
      }
    }

    const data = dataSheet.getDataRange().getDisplayValues();
    let dumpIdData = findCol(dumpRawTop, "fuze", "id");
    let matchIdx = dumpIdData.col; let dumpHeaderRow = dumpIdData.row;
    let latIdx = findCol(dumpRawTop, "latitude").col; let lngIdx = findCol(dumpRawTop, "longitude").col;
    let nameIdx = findCol(dumpRawTop, "site name").col;
    let headers = data[dumpHeaderRow]; let parsedRows = [];
    
    function hasDate(val) {
      if (!val) return false;
      let s = String(val).toLowerCase().trim();
      if (s === "" || s === "tbd" || s === "n/a" || s === "tbc" || s === "pending") return false;
      return /\d/.test(s);
    }
    
    for (let i = dumpHeaderRow + 1; i < data.length; i++) {
      let currentId = String(data[i][matchIdx]).toLowerCase().replace(/\D/g, "");
      if (validSites.has(currentId)) {
        let siteData = validSites.get(currentId);
        
        let milestoneStatus = "Pending";
        let milestoneColor = "#ff4d4d"; 
        
        if (actIdx > -1 && hasDate(data[i][actIdx])) { 
          milestoneStatus = "Activated"; milestoneColor = "#36ab1f"; 
        } else if (intIdx > -1 && hasDate(data[i][intIdx])) { 
          milestoneStatus = "Integrated"; milestoneColor = "#fbbc04"; 
        } else if (rdyIdx > -1 && hasDate(data[i][rdyIdx])) { 
          milestoneStatus = "BB Ready"; milestoneColor = "#f4cccc"; 
        }

        parsedRows.push({ 
          values: data[i], 
          color: siteData.color,
          milestoneColor: milestoneColor, 
          milestoneStatus: milestoneStatus,
          clusterId: siteData.clusterId,
          hub: siteData.hub,
          displayBbu: siteData.cleanBbu
        });
      }
    }

    return JSON.stringify({ headers: headers, rows: parsedRows, latIdx: latIdx, lngIdx: lngIdx, nameIdx: nameIdx, hubCoords: frontendHubCoords });

  } catch (e) {
    return JSON.stringify({ error: e.toString() });
  }
}

function doGet() {
  const html = `
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
      <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
      <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.4.1/dist/MarkerCluster.css" />
      <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.4.1/dist/MarkerCluster.Default.css" />
      <script src="https://unpkg.com/leaflet.markercluster@1.4.1/dist/leaflet.markercluster.js"></script>
      <style>
        html, body { height: 100%; width: 100%; margin: 0; padding: 0; background-color: #e5e5e5; font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; display: flex; flex-direction: row; overflow: hidden; }
        #app-container { display: flex; width: 100%; height: 100vh; }
        #sidebar { width: 320px; flex-shrink: 0; background: white; border-right: 1px solid #ccc; display: flex; flex-direction: column; z-index: 1000; box-shadow: 2px 0 10px rgba(0,0,0,0.1); }
        #sidebar-header { padding: 25px 15px 20px 15px; border-bottom: 1px solid #eee; background: #fff; }
        #sidebar-header h3 { margin: 0 0 5px 0; font-size: 19px; font-weight: 700; color: #000; letter-spacing: -0.5px; }
        .contact-info { font-size: 11px; color: #666; margin-bottom: 20px; line-height: 1.4; }
        #search-box { display: flex; width: 100%; margin-bottom: 15px; }
        #search-input { flex-grow: 1; padding: 10px; border: 1px solid #d8d8d8; border-radius: 2px 0 0 2px; font-size: 14px; outline: none; }
        #search-btn { padding: 10px 15px; background: #ed0000; color: white; border: none; cursor: pointer; font-weight: bold; font-size: 13px; text-transform: uppercase; }
        .filter-buttons { display: flex; gap: 8px; margin-bottom: 5px; flex-wrap: wrap;}
        .filter-buttons button { flex: 1 1 45%; padding: 6px; cursor: pointer; border: 1px solid #ccc; border-radius: 2px; background: #f9f9f9; font-size: 11px; font-weight: 600; }
        .filter-buttons button.mode-btn { flex: 1 1 100%; background: #000; color: #fff; border-color: #000; margin-top: 5px; }
        .filter-buttons button.hub-toggle-btn { flex: 1 1 100%; background: #555; color: #fff; border-color: #555; margin-top: 5px; }
        #hub-list { flex-grow: 1; overflow-y: auto; padding: 10px 5px; }
        .hub-row { display: flex; align-items: center; padding: 8px 10px; cursor: pointer; border-radius: 4px; transition: background 0.2s ease; user-select: none; }
        .hub-row input { pointer-events: none; margin-right: 12px; transform: scale(1.1); }
        .hub-row span { font-size: 13px; color: #333; font-weight: 500; }
        #map { flex-grow: 1; height: 100%; width: 100%; background-color: #ddd; position: relative; }
        #loading-overlay { position: fixed; top: 0; left: 0; width: 100vw; height: 100vh; background: #fff; z-index: 9999; display: flex; flex-direction: column; align-items: center; justify-content: center; }
        .progress-container { width: 280px; background-color: #eee; border-radius: 10px; overflow: hidden; margin-top: 20px; }
        .progress-bar { height: 12px; width: 0%; background-color: #ed0000; transition: width 0.1s; }
        .site-label { background: transparent !important; border: none !important; box-shadow: none !important; font-weight: bold; font-size: 11px; color: #000; text-shadow: 1px 1px 0 #fff, -1px -1px 0 #fff, 1px -1px 0 #fff, -1px 1px 0 #fff; }
        .hub-master-label { background: transparent !important; border: none !important; box-shadow: none !important; font-weight: 900; font-size: 14px; color: #cc0000; text-shadow: 2px 2px 0 #fff, -2px -2px 0 #fff, 2px -2px 0 #fff, -2px 2px 0 #fff; }
        .popup-content { font-size: 12px; max-height: 250px; overflow-y: auto; min-width: 220px; }
        .popup-content ul { padding-left: 15px; margin: 5px 0; list-style-type: square; }
        .debug-id { font-family: monospace; background: #f0f0f0; padding: 2px 4px; border-radius: 3px; color: #d00; font-size: 10px;}
        #map-legend { background: rgba(255, 255, 255, 0.95); padding: 15px; border-radius: 8px; box-shadow: 0 4px 15px rgba(0,0,0,0.15); font-size: 12px; color: #333; line-height: 1.8; min-width: 140px; }
      </style>
    </head>
    <body>
      <div id="loading-overlay">
        <div id="progress-text" style="font-size:16px; font-weight:bold; color:#000;">Initializing Network Data...</div>
        <div class="progress-container"><div id="progress-bar-fill" class="progress-bar"></div></div>
      </div>
      <div id="app-container">
        <div id="sidebar">
          <div id="sidebar-header">
            <h3>Small Cell BBU Cluster Map</h3>
            <div class="contact-info">Reach out to Sean McGrath or Enis Orahovac with any questions</div>
            <div id="search-box">
              <input type="text" id="search-input" placeholder="Find Site or Project ID..." />
              <button id="search-btn" onclick="executeSearch()">Search</button>
            </div>
            <div class="filter-buttons">
              <button onclick="toggleAllHubs(true)">Select All</button>
              <button onclick="toggleAllHubs(false)">Clear Map</button>
              <button id="hub-toggle-btn" class="hub-toggle-btn" onclick="toggleHubPins()">📍 Toggle Hub Pins (On/Off)</button>
              <button id="mode-btn" class="mode-btn" onclick="toggleMode()">👁️ Switch to Milestone Mode</button>
            </div>
          </div>
          <div id="hub-list"></div>
        </div>
        <div id="map"></div>
      </div>

      <script>
        const map = L.map('map', { tap: false, zoomControl: false }).setView([39.82, -98.57], 4);
        L.control.zoom({ position: 'bottomright' }).addTo(map);
        L.tileLayer('https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png', { attribution: '© CARTO' }).addTo(map);
        const markers = L.markerClusterGroup({ chunkedLoading: true, disableClusteringAtZoom: 17 });
        map.addLayer(markers);
        
        const legend = L.control({ position: 'bottomleft' });
        legend.onAdd = function () {
          const div = L.DomUtil.create('div', 'info legend');
          div.id = 'map-legend';
          return div;
        };
        legend.addTo(map);
        
        window.siteDataMap = []; 
        let hubGroups = {}; 
        let activePolylines = []; 
        let activeHubMarkers = []; 
        let isMilestoneMode = false;
        let showHubPins = true;

        const progressText = document.getElementById('progress-text');
        const progressBarFill = document.getElementById('progress-bar-fill');

        google.script.run
          .withSuccessHandler(function(responseString) {
            const data = JSON.parse(responseString);
            if(data.error) { alert("Data Error: " + data.error); return; }
            
            const { headers, rows, latIdx, lngIdx, nameIdx, hubCoords } = data; 
            let currentIndex = 0; const chunkSize = 200; 

            function processChunk() {
              const end = Math.min(currentIndex + chunkSize, rows.length);
              for (let i = currentIndex; i < end; i++) {
                const rowObj = rows[i];
                const row = rowObj.values;
                const lat = parseFloat(row[latIdx]), lng = parseFloat(row[lngIdx]);
                
                if(!isNaN(lat) && !isNaN(lng)) {
                  let title = row[nameIdx] || 'Unnamed';
                  let popupHtml = '<div class="popup-content"><b>' + title + '</b><hr><ul>';
                  popupHtml += '<li><b>Map Cluster ID:</b> <span class="debug-id">' + rowObj.clusterId + '</span></li>';
                  popupHtml += '<li><b>BBU Link:</b> ' + (rowObj.displayBbu || 'None') + '</li>';
                  popupHtml += '<li><b>Status:</b> <strong style="color:'+rowObj.milestoneColor+'">' + rowObj.milestoneStatus + '</strong></li>';
                  headers.forEach((h, j) => { if(row[j] && row[j].toString().trim() !== "") popupHtml += '<li><b>' + h + ':</b> ' + row[j] + '</li>'; });
                  popupHtml += '</ul></div>';

                  const marker = L.marker([lat, lng])
                    .bindTooltip(title, { permanent: true, direction: 'right', className: 'site-label', offset: [8, 0], interactive: true })
                    .bindPopup(popupHtml);
                  
                  marker.clusterColor = rowObj.color;
                  marker.milestoneColor = rowObj.milestoneColor;
                  
                  marker.setIcon(L.divIcon({ className: 'custom-div-icon', html: "<div style='background-color:" + marker.clusterColor + "; width: 14px; height: 14px; border-radius: 50%; border: 2px solid white; box-shadow: 0 0 3px rgba(0,0,0,0.5);'></div>", iconSize: [18, 18], iconAnchor: [9, 9] }));
                  marker.on('click', () => marker.openPopup());
                  
                  let hub = rowObj.hub; 
                  if (!hubGroups[hub]) hubGroups[hub] = { markers: [], clusters: {}, hubMarkers: [] };
                  hubGroups[hub].markers.push(marker);

                  let cId = rowObj.clusterId;
                  if (!hubGroups[hub].clusters[cId]) hubGroups[hub].clusters[cId] = { points: [], color: rowObj.color };
                  hubGroups[hub].clusters[cId].points.push([lat, lng]);

                  window.siteDataMap.push({ marker: marker, name: title, hub: hub, searchText: (title + " " + row.join(" ")).toLowerCase() });
                }
              }
              currentIndex = end;
              progressBarFill.style.width = Math.round((currentIndex / rows.length) * 100) + '%';
              if (currentIndex < rows.length) { setTimeout(processChunk, 10); } else { finishLoading(hubCoords); }
            }
            processChunk();
          }).getMapData();

        function finishLoading(hubCoords) {
          for (let hub in hubGroups) {
            hubGroups[hub].polylines = [];
            if (hubCoords && hubCoords[hub]) {
              
              let isIen = hubCoords[hub].isIen;
              let hubTypeLabel = isIen ? "iEN Hub" : "Non-iEN Hub";
              
              // Gold Star for iEN
              const ienIcon = L.divIcon({
                className: 'custom-hub-icon',
                html: "<div style='background-color:#1a1a1a; width: 22px; height: 22px; border-radius: 6px; border: 2px solid #FFD700; display:flex; justify-content:center; align-items:center; color:#FFD700; font-size:16px; box-shadow: 0 4px 6px rgba(0,0,0,0.5);'>★</div>",
                iconSize: [26, 26], iconAnchor: [13, 13]
              });
              
              // Blue Diamond for Non-iEN
              const nonIenIcon = L.divIcon({
                className: 'custom-hub-icon',
                html: "<div style='background-color:#1a1a1a; width: 22px; height: 22px; border-radius: 6px; border: 2px solid #00BFFF; display:flex; justify-content:center; align-items:center; color:#00BFFF; font-size:16px; box-shadow: 0 4px 6px rgba(0,0,0,0.5);'>♦</div>",
                iconSize: [26, 26], iconAnchor: [13, 13]
              });
              
              const masterHubIcon = isIen ? ienIcon : nonIenIcon;
              
              const hMarker = L.marker([hubCoords[hub].lat, hubCoords[hub].lng], { icon: masterHubIcon, zIndexOffset: 1000 })
                .bindTooltip(hubTypeLabel + ": " + hub, { permanent: true, direction: 'top', className: 'hub-master-label', offset: [0, -10] })
                .bindPopup("<div class='popup-content'><h3 style='margin:0;color:#cc0000;'>" + hub + "</h3><p style='margin-bottom:0;'>" + hubTypeLabel + "</p></div>");
              
              hubGroups[hub].hubMarkers.push(hMarker);
            }
            for (let cId in hubGroups[hub].clusters) {
              let cluster = hubGroups[hub].clusters[cId];
              if (cluster.points.length > 1) {
                if (cluster.points.length > 2) cluster.points.push(cluster.points[0]); 
                let line = L.polyline(cluster.points, { color: cluster.color, weight: 4, opacity: 0.8 }).bindPopup("Cluster ID: " + cId);
                line.originalColor = cluster.color; 
                hubGroups[hub].polylines.push(line);
              }
            }
          }

          Object.keys(hubGroups).sort().forEach(h => {
            if (!h) return;
            const div = document.createElement('div');
            let safeVal = String(h).replace(/"/g, '&quot;'); 
            div.innerHTML = '<label class="hub-row"><input type="checkbox" class="hub-cb" value="'+safeVal+'" checked> <span>'+h+'</span></label>';
            document.getElementById('hub-list').appendChild(div);
          });
          
          document.querySelectorAll('.hub-cb').forEach(cb => cb.addEventListener('change', updateMap));
          updateMap();
          document.getElementById('loading-overlay').style.display = 'none';
          setTimeout(() => { map.invalidateSize(); if (markers.getLayers().length > 0) map.fitBounds(markers.getBounds(), {padding:[40,40]}); }, 200);
        }

        function updateLegend() {
          const lgd = document.getElementById('map-legend');
          if (!lgd) return;
          if (isMilestoneMode) {
            lgd.innerHTML = '<strong style="font-size:14px; margin-bottom:6px; display:block; color:#000;">Schedule Status</strong>' +
              '<div><span style="display:inline-block; width:12px; height:12px; background:#36ab1f; border-radius:50%; margin-right:8px; border:1px solid #fff; box-shadow:0 0 2px rgba(0,0,0,0.5);"></span> Activated</div>' +
              '<div><span style="display:inline-block; width:12px; height:12px; background:#fbbc04; border-radius:50%; margin-right:8px; border:1px solid #fff; box-shadow:0 0 2px rgba(0,0,0,0.5);"></span> Integrated</div>' +
              '<div><span style="display:inline-block; width:12px; height:12px; background:#f4cccc; border-radius:50%; margin-right:8px; border:1px solid #fff; box-shadow:0 0 2px rgba(0,0,0,0.5);"></span> BB Ready</div>' +
              '<div><span style="display:inline-block; width:12px; height:12px; background:#ff4d4d; border-radius:50%; margin-right:8px; border:1px solid #fff; box-shadow:0 0 2px rgba(0,0,0,0.5);"></span> Pending</div>' +
              '<hr style="margin:8px 0; border:0; border-top:1px solid #ccc;">' +
              '<div><span style="display:inline-block; width:14px; height:0; border-top:2px dashed #a0a0a0; margin-right:6px; vertical-align:middle;"></span> Cluster Link</div>' +
              '<div><span style="display:inline-flex; justify-content:center; align-items:center; width:16px; height:16px; background:#1a1a1a; color:#FFD700; border-radius:4px; margin-right:6px; font-size:10px; border:1px solid #FFD700;">★</span> iEN Hub</div>' +
              '<div><span style="display:inline-flex; justify-content:center; align-items:center; width:16px; height:16px; background:#1a1a1a; color:#00BFFF; border-radius:4px; margin-right:6px; font-size:10px; border:1px solid #00BFFF;">♦</span> Non-iEN Hub</div>';
          } else {
            lgd.innerHTML = '<strong style="font-size:14px; margin-bottom:6px; display:block; color:#000;">BBU Cluster Map</strong>' +
              '<div><span style="display:inline-block; width:12px; height:12px; background:#3388ff; border-radius:50%; margin-right:8px; border:1px solid #fff; box-shadow:0 0 2px rgba(0,0,0,0.5);"></span> Paired Node</div>' +
              '<div><span style="display:inline-block; width:14px; height:4px; background:#3388ff; margin-right:6px; vertical-align:middle; border-radius:2px;"></span> Pairing Link</div>' +
              '<hr style="margin:8px 0; border:0; border-top:1px solid #ccc;">' +
              '<div><span style="display:inline-flex; justify-content:center; align-items:center; width:16px; height:16px; background:#1a1a1a; color:#FFD700; border-radius:4px; margin-right:6px; font-size:10px; border:1px solid #FFD700;">★</span> iEN Hub</div>' +
              '<div><span style="display:inline-flex; justify-content:center; align-items:center; width:16px; height:16px; background:#1a1a1a; color:#00BFFF; border-radius:4px; margin-right:6px; font-size:10px; border:1px solid #00BFFF;">♦</span> Non-iEN Hub</div>';
          }
        }

        function updateMap() {
          markers.clearLayers();
          activePolylines.forEach(l => map.removeLayer(l));
          activeHubMarkers.forEach(m => map.removeLayer(m)); 
          activePolylines = []; activeHubMarkers = []; 
          
          document.querySelectorAll('.hub-cb:checked').forEach(cb => {
            const group = hubGroups[cb.value];
            if (group) {
              markers.addLayers(group.markers); 
              
              if (showHubPins && group.hubMarkers) {
                group.hubMarkers.forEach(m => { map.addLayer(m); activeHubMarkers.push(m); });
              }
              
              group.polylines.forEach(l => { 
                if (isMilestoneMode) {
                  l.setStyle({color: '#a0a0a0', weight: 2, opacity: 0.6, dashArray: '5, 5'});
                } else {
                  l.setStyle({color: l.originalColor, weight: 4, opacity: 0.8, dashArray: null});
                }
                map.addLayer(l); 
                activePolylines.push(l); 
              });
            }
          });
          updateLegend();
        }

        function toggleHubPins() {
          showHubPins = !showHubPins;
          const btn = document.getElementById('hub-toggle-btn');
          if (showHubPins) {
            btn.style.backgroundColor = "#555"; 
            btn.style.borderColor = "#555";
          } else {
            btn.style.backgroundColor = "#999";
            btn.style.borderColor = "#999";
          }
          updateMap();
        }

        function toggleMode() {
          isMilestoneMode = !isMilestoneMode;
          const btn = document.getElementById('mode-btn');
          
          if (isMilestoneMode) {
            btn.innerHTML = "👁️ Switch to Cluster Mode";
            btn.style.backgroundColor = "#d52b1e"; 
            btn.style.borderColor = "#d52b1e";
          } else {
            btn.innerHTML = "👁️ Switch to Milestone Mode";
            btn.style.backgroundColor = "#000";
            btn.style.borderColor = "#000";
          }
          
          window.siteDataMap.forEach(site => {
            let activeColor = isMilestoneMode ? site.marker.milestoneColor : site.marker.clusterColor;
            site.marker.setIcon(L.divIcon({ className: 'custom-div-icon', html: "<div style='background-color:" + activeColor + "; width: 14px; height: 14px; border-radius: 50%; border: 2px solid white; box-shadow: 0 0 3px rgba(0,0,0,0.5);'></div>", iconSize: [18, 18], iconAnchor: [9, 9] }));
          });
          updateMap();
        }

        function toggleAllHubs(v) { document.querySelectorAll('.hub-cb').forEach(cb => cb.checked = v); updateMap(); }
        
        function executeSearch() {
          const q = document.getElementById('search-input').value.toLowerCase();
          const match = window.siteDataMap.find(s => s.searchText.includes(q));
          if (match) {
            let safeVal = match.hub.replace(/"/g, '&quot;');
            const cb = document.querySelector('.hub-cb[value="'+safeVal+'"]');
            if (cb && !cb.checked) { cb.checked = true; updateMap(); }
            markers.zoomToShowLayer(match.marker, () => match.marker.openPopup());
          } else { alert("No results found."); }
        }
      </script>
    </body>
    </html>
  `;
  return HtmlService.createHtmlOutput(html).setTitle('Small Cell BBU Cluster Map').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}