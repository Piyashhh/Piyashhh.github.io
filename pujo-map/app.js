// ========================================
// PUJO MAP — VERSION 0.3
// ========================================


// ----------------------------------------
// 1. MAP
// ----------------------------------------

const map = L.map("map").setView(
    [22.5726, 88.3639],
    12
);


L.tileLayer(
    "https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png",
    {
        attribution: "&copy; OpenStreetMap contributors"
    }
).addTo(map);


// ----------------------------------------
// 2. APPLICATION STATE
// ----------------------------------------

let pandals = [];

let markers = [];

let plan = [];



let routeLayer = null;

let draggedPlanIndex = null;


// ----------------------------------------
// 3. LOAD SAVED PLAN
// ----------------------------------------

function loadSavedPlan() {

    const savedPlan =
        localStorage.getItem("pujoPlan");

    if (!savedPlan) {
        plan = [];
        return;
    }

    try {

        plan = JSON.parse(savedPlan);

    } catch (error) {

        console.error(
            "Could not load saved plan:",
            error
        );

        plan = [];
    }
}


// ----------------------------------------
// 4. SAVE PLAN
// ----------------------------------------

function savePlan() {

    localStorage.setItem(
        "pujoPlan",
        JSON.stringify(plan)
    );

    updatePlanUI();
}


// ----------------------------------------
// 5. LOAD EXCEL
// ----------------------------------------

async function loadPandals() {

    try {

        const response = await fetch(
            "data/pandals.xlsx?v=" + Date.now()
        );

        if (!response.ok) {
            throw new Error(
                "Could not load pandals.xlsx"
            );
        }

        const fileData =
            await response.arrayBuffer();

        const workbook =
            XLSX.read(
                fileData,
                {
                    type: "array"
                }
            );

        const sheetName =
            workbook.SheetNames[0];

        const sheet =
            workbook.Sheets[sheetName];

        pandals =
            XLSX.utils.sheet_to_json(sheet);


        console.log(
            "TOTAL PANDALS:",
            pandals.length
        );


        displayPandals(pandals);

        updatePlanUI();

    } catch (error) {

        console.error(error);

        document.getElementById(
            "pandal-count"
        ).textContent =
            "Could not load data";
    }
}


// ----------------------------------------
// 6. DISPLAY PANDALS
// ----------------------------------------

function displayPandals(data) {

    markers.forEach(marker => {
        map.removeLayer(marker);
    });

    markers = [];

    let validCount = 0;


    data.forEach(pandal => {

        const lat =
            Number(pandal["Latitude"]);

        const lng =
            Number(pandal["Longitude"]);


        if (
            !Number.isFinite(lat) ||
            !Number.isFinite(lng)
        ) {
            return;
        }


        validCount++;


        const marker =
            L.marker([
                lat,
                lng
            ]);


        marker.pandalData = pandal;


        marker.bindPopup(
            createPandalPopup(pandal)
        );


        marker.addTo(map);


        markers.push(marker);

    });


    document.getElementById(
        "pandal-count"
    ).textContent =
        `${validCount} pandals`;


    console.log(
        "Markers created:",
        markers.length
    );


    setTimeout(() => {
        map.invalidateSize();
    }, 100);
}


// ----------------------------------------
// 7. CREATE PANDAL POPUP
// ----------------------------------------

function createPandalPopup(pandal) {

    const name =
        escapeHTML(
            pandal["Puja Name"] ||
            "Unnamed Puja"
        );

    const address =
        escapeHTML(
            pandal["Address"] ||
            "Address unavailable"
        );

    const id =
        escapeHTML(
            pandal["Pandal ID"] ||
            ""
        );

    const mapsUrl =
        pandal["Google Maps Link"] ||
        "#";


    const alreadyAdded =
        isInPlan(pandal);


    const addButtonText =
        alreadyAdded
            ? "✓ In My Plan"
            : "+ Add to Plan";


    return `
        <div class="pandal-popup">

            <h3>${name}</h3>

            <span class="pandal-id">
                ${id}
            </span>

            <div class="pandal-address">
                ${address}
            </div>

            <div class="popup-buttons">

                <a
                    class="maps-button"
                    href="${mapsUrl}"
                    target="_blank"
                    rel="noopener noreferrer"
                >
                    Navigate
                </a>

                <button
                    class="add-plan-button"
                    onclick="togglePlan('${escapeAttribute(id)}')"
                >
                    ${addButtonText}
                </button>

            </div>

        </div>
    `;
}


// ----------------------------------------
// 8. CHECK WHETHER PANDAL IS IN PLAN
// ----------------------------------------

function isInPlan(pandal) {

    const id =
        String(pandal["Pandal ID"]);


    return plan.some(
        item =>
            String(item["Pandal ID"]) === id
    );
}


// ----------------------------------------
// 9. ADD / REMOVE PANDAL
// ----------------------------------------

function togglePlan(id) {

    const index =
        plan.findIndex(
            item =>
                String(item["Pandal ID"]) ===
                String(id)
        );


    // Already in plan → remove
    if (index !== -1) {

        plan.splice(index, 1);

    }

    // Not in plan → add
    else {

        const pandal =
            pandals.find(
                item =>
                    String(item["Pandal ID"]) ===
                    String(id)
            );


        if (!pandal) {
            return;
        }


        plan.push(pandal);
    }


    savePlan();


    // Refresh popup button
    const marker =
        markers.find(
            marker =>
                String(
                    marker.pandalData["Pandal ID"]
                ) === String(id)
        );


    if (marker) {

        marker.setPopupContent(
            createPandalPopup(
                marker.pandalData
            )
        );

        marker.openPopup();
    }
}


// ----------------------------------------
// 10. UPDATE PLAN UI
// ----------------------------------------

function updatePlanUI() {

    const count =
        plan.length;


    document.getElementById(
        "plan-count"
    ).textContent =
        count;


    document.getElementById(
        "plan-summary"
    ).textContent =
        `${count} ${count === 1 ? "pandal" : "pandals"}`;


    const list =
        document.getElementById(
            "plan-list"
        );


    list.innerHTML = "";


    if (count === 0) {

        list.innerHTML = `
            <div class="plan-empty">

                Your plan is empty.

                <br>

                Click a pandal and add it.

            </div>
        `;

        return;
    }


    plan.forEach(
    (pandal, index) => {

        const item =
            document.createElement("div");

        item.className =
            "plan-item";

        item.draggable = true;

        item.dataset.index = index;

        item.innerHTML = `

                <div class="plan-number">
                    ${index + 1}
                </div>

                <div class="plan-info">

                    <div class="plan-name">
                        ${escapeHTML(
                            pandal["Puja Name"] ||
                            "Unnamed Puja"
                        )}
                    </div>

                    <div class="plan-address">
                        ${escapeHTML(
                            pandal["Address"] ||
                            ""
                        )}
                    </div>

                </div>

                <button
                    class="remove-plan-item"
                    data-id="${escapeAttribute(
                        pandal["Pandal ID"]
                    )}"
                >
                    ×
                </button>
            `;

            item.addEventListener(
    "dragstart",
    () => {

        draggedPlanIndex = index;

        item.classList.add("dragging");
    }
);

item.addEventListener(
    "dragend",
    () => {

        draggedPlanIndex = null;

        item.classList.remove("dragging");

        document
            .querySelectorAll(".plan-item")
            .forEach(
                element =>
                    element.classList.remove(
                        "drag-over"
                    )
            );
    }
);

item.addEventListener(
    "dragover",
    event => {

        event.preventDefault();

        item.classList.add("drag-over");
    }
);

item.addEventListener(
    "dragleave",
    () => {

        item.classList.remove("drag-over");
    }
);

item.addEventListener(
    "drop",
    event => {

        event.preventDefault();

        item.classList.remove("drag-over");

        const targetIndex =
            Number(item.dataset.index);

        if (
            draggedPlanIndex === null ||
            draggedPlanIndex === targetIndex
        ) {
            return;
        }

        const movedPandal =
            plan.splice(
                draggedPlanIndex,
                1
            )[0];

        plan.splice(
            targetIndex,
            0,
            movedPandal
        );

        savePlan();

        clearRoute();
    }
);

            // Clicking the item focuses its marker
            item.addEventListener(
                "click",
                event => {

                    if (
                        event.target.classList.contains(
                            "remove-plan-item"
                        )
                    ) {
                        return;
                    }

                    focusPandal(pandal);
                }
            );


            // Remove button
            item.querySelector(
                ".remove-plan-item"
            ).addEventListener(
                "click",
                () => {

                    togglePlan(
                        pandal["Pandal ID"]
                    );

                }
            );


            list.appendChild(item);

        }
    );
}


// ----------------------------------------
// 11. FOCUS PANDAL
// ----------------------------------------

function focusPandal(pandal) {

    const lat =
        Number(pandal["Latitude"]);

    const lng =
        Number(pandal["Longitude"]);


    if (
        !Number.isFinite(lat) ||
        !Number.isFinite(lng)
    ) {
        return;
    }


    map.setView(
        [lat, lng],
        17,
        {
            animate: true
        }
    );


    const marker =
        markers.find(
            marker =>
                String(
                    marker.pandalData["Pandal ID"]
                ) ===
                String(
                    pandal["Pandal ID"]
                )
        );


    if (marker) {
        marker.openPopup();
    }


    closePlanPanel();
}


// ----------------------------------------
// 12. PLAN PANEL
// ----------------------------------------

const planButton =
    document.getElementById(
        "plan-button"
    );

const planPanel =
    document.getElementById(
        "plan-panel"
    );

const closePlan =
    document.getElementById(
        "close-plan"
    );


planButton.addEventListener(
    "click",
    () => {

        planPanel.style.display =
            "flex";

    }
);


closePlan.addEventListener(
    "click",
    closePlanPanel
);


function closePlanPanel() {

    planPanel.style.display =
        "none";
}


// ----------------------------------------
// 13. CLEAR PLAN
// ----------------------------------------

document.getElementById(
    "clear-plan"
).addEventListener(
    "click",
    () => {

        if (plan.length === 0) {
            return;
        }


        const confirmed =
            confirm(
                "Clear your entire Pujo plan?"
            );


        if (!confirmed) {
            return;
        }


        plan = [];

        savePlan();


        // Refresh all marker popups
        markers.forEach(
            marker => {

                marker.setPopupContent(
                    createPandalPopup(
                        marker.pandalData
                    )
                );

            }
        );

    }
);


// ----------------------------------------
// 14. SEARCH
// ----------------------------------------

const searchInput =
    document.getElementById(
        "search"
    );

const searchResults =
    document.getElementById(
        "search-results"
    );

const clearButton =
    document.getElementById(
        "clear-search"
    );


searchInput.addEventListener(
    "input",
    handleSearch
);


function handleSearch() {

    const query =
        searchInput.value
            .trim()
            .toLowerCase();


    clearButton.style.display =
        query
            ? "block"
            : "none";


    if (!query) {

        searchResults.style.display =
            "none";

        showAllMarkers();

        return;
    }


    const results =
        pandals.filter(
            pandal => {

                const name =
                    String(
                        pandal["Puja Name"] ||
                        ""
                    ).toLowerCase();


                const address =
                    String(
                        pandal["Address"] ||
                        ""
                    ).toLowerCase();


                return (
                    name.includes(query) ||
                    address.includes(query)
                );

            }
        );


    displaySearchResults(results);

    showMatchingMarkers(results);
}


// ----------------------------------------
// 15. SEARCH RESULTS
// ----------------------------------------

function displaySearchResults(results) {

    searchResults.innerHTML = "";


    if (results.length === 0) {

        searchResults.innerHTML = `
            <div class="no-results">
                No pandals found.
            </div>
        `;

        searchResults.style.display =
            "block";

        return;
    }


    results
        .slice(0, 10)
        .forEach(pandal => {

            const result =
                document.createElement(
                    "div"
                );


            result.className =
                "search-result";


            result.innerHTML = `

                <div class="search-result-name">

                    ${escapeHTML(
                        pandal["Puja Name"] ||
                        "Unnamed Puja"
                    )}

                </div>

                <div class="search-result-address">

                    ${escapeHTML(
                        pandal["Address"] ||
                        ""
                    )}

                </div>
            `;


            result.addEventListener(
                "click",
                () => selectPandal(pandal)
            );


            searchResults.appendChild(
                result
            );

        });


    searchResults.style.display =
        "block";
}


// ----------------------------------------
// 16. SELECT SEARCH RESULT
// ----------------------------------------

function selectPandal(pandal) {

    focusPandal(pandal);

    searchResults.style.display =
        "none";
}


// ----------------------------------------
// 17. FILTER MARKERS
// ----------------------------------------

function showMatchingMarkers(results) {

    const resultIDs =
        new Set(
            results.map(
                pandal =>
                    String(
                        pandal["Pandal ID"]
                    )
            )
        );


    markers.forEach(
        marker => {

            const id =
                String(
                    marker.pandalData["Pandal ID"]
                );


            if (resultIDs.has(id)) {

                marker.addTo(map);

            } else {

                map.removeLayer(marker);

            }

        }
    );
}


// ----------------------------------------
// 18. SHOW ALL MARKERS
// ----------------------------------------

function showAllMarkers() {

    markers.forEach(
        marker => {

            marker.addTo(map);

        }
    );
}


// ----------------------------------------
// 19. CLEAR SEARCH
// ----------------------------------------

clearButton.addEventListener(
    "click",
    () => {

        searchInput.value = "";

        searchResults.style.display =
            "none";

        clearButton.style.display =
            "none";

        showAllMarkers();

        searchInput.focus();

    }
);


// ----------------------------------------
// 20. CLOSE SEARCH
// ----------------------------------------

map.on(
    "click",
    () => {

        searchResults.style.display =
            "none";

    }
);


// ----------------------------------------
// 21. HTML ESCAPING
// ----------------------------------------

function escapeHTML(value) {

    return String(value)
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}


function escapeAttribute(value) {

    return String(value)
        .replaceAll("\\", "\\\\")
        .replaceAll("'", "\\'");
}

// ----------------------------------------
// ROUTING
// ----------------------------------------

async function showRoute() {

    if (plan.length < 2) {

        alert(
            "Add at least 2 pandals to create a route."
        );

        return;
    }


    const coordinates = plan
        .map(pandal => {

            const lat =
                Number(pandal["Latitude"]);

            const lng =
                Number(pandal["Longitude"]);

            return `${lng},${lat}`;
        })
        .join(";");


    const url =
        `https://router.project-osrm.org/route/v1/driving/` +
        `${coordinates}?overview=full&geometries=geojson`;


    const button =
        document.getElementById("show-route");


    button.textContent =
        "Calculating...";

    button.disabled = true;


    try {

        const response =
            await fetch(url);


        if (!response.ok) {
            throw new Error(
                "Routing request failed."
            );
        }


        const data =
            await response.json();


        if (
            data.code !== "Ok" ||
            !data.routes ||
            !data.routes.length
        ) {
            throw new Error(
                "No route found."
            );
        }


        const route =
            data.routes[0];


        drawRoute(
            route.geometry
        );


        displayRouteSummary(
            route
        );


    } catch (error) {

        console.error(error);

        alert(
            "Couldn't calculate the route. Try again."
        );

    } finally {

        button.textContent =
            "Show Route";

        button.disabled = false;
    }
}

// ----------------------------------------
// DRAW ROUTE
// ----------------------------------------

function drawRoute(geometry) {

    // Remove previous route
    if (routeLayer) {

        map.removeLayer(
            routeLayer
        );

    }


    routeLayer =
        L.geoJSON(
            geometry,
            {
                style: {
                    weight: 5,
                    opacity: 0.8
                }
            }
        ).addTo(map);


    // Fit map around route
    map.fitBounds(
        routeLayer.getBounds(),
        {
            padding: [50, 50]
        }
    );
}

// ----------------------------------------
// ROUTE SUMMARY
// ----------------------------------------

function displayRouteSummary(route) {

    const summary =
        document.getElementById(
            "route-summary"
        );


    const distanceKm =
        (route.distance / 1000)
            .toFixed(1);


    const minutes =
        Math.round(
            route.duration / 60
        );


    let timeText;


    if (minutes < 60) {

        timeText =
            `${minutes} min`;

    } else {

        const hours =
            Math.floor(minutes / 60);

        const remaining =
            minutes % 60;

        timeText =
            `${hours}h ${remaining}m`;
    }


    summary.innerHTML = `
        <strong>${distanceKm} km</strong>
        · approximately
        <strong>${timeText}</strong>
        by road
    `;


    summary.style.display =
        "block";
}

document.getElementById(
    "show-route"
).addEventListener(
    "click",
    showRoute
);

// ----------------------------------------
// CLEAR ROUTE
// ----------------------------------------

function clearRoute() {

    if (routeLayer) {

        map.removeLayer(
            routeLayer
        );

        routeLayer = null;
    }


    const summary =
        document.getElementById(
            "route-summary"
        );


    summary.style.display =
        "none";


    summary.innerHTML =
        "";
}

// ----------------------------------------
// 22. START
// ----------------------------------------

loadSavedPlan();

loadPandals();