import * as THREE from "three";
import { OrbitControls } from "three/addons/controls/OrbitControls.js";
import { RoomEnvironment } from "three/addons/environments/RoomEnvironment.js";

const MM_TO_M = 0.001;
const PALLET_SELF_WEIGHT_KG = 20;
const TOUCH_TOL = 0.05;

const DEFAULT_CONTAINERS = {
  "20GP": { label: "20' GP", lengthMm: 5890, widthMm: 2300, heightMm: 2390, tareKg: 2250, maxGrossKg: 30480, maxPayloadKg: 28230 },
  "40GP": { label: "40' GP", lengthMm: 11960, widthMm: 2300, heightMm: 2390, tareKg: 3780, maxGrossKg: 30480, maxPayloadKg: 26700 },
  "40HQ": { label: "40' HQ", lengthMm: 11960, widthMm: 2300, heightMm: 2570, tareKg: 3980, maxGrossKg: 30480, maxPayloadKg: 26500 },
};
let CONTAINERS = { ...DEFAULT_CONTAINERS };

const PALLET_MODELS = [
  { id: "inv25-A", product: "10-25kW Hybrid Inverter", model: "A", units: 6, size: [1170, 730, 1480], unitWeight: 60, color: "#ff4d9e" },
  { id: "inv25-B", product: "10-25kW Hybrid Inverter", model: "B", units: 4, size: [1170, 730, 1030], unitWeight: 60, color: "#ff6f00" },
  { id: "inv50-A", product: "30-50kW Hybrid Inverter", model: "A", units: 3, size: [1030, 800, 1480], unitWeight: 100, color: "#ffbf00" },
  { id: "inv50-B", product: "30-50kW Hybrid Inverter", model: "B", units: 2, size: [1030, 800, 1030], unitWeight: 100, color: "#53d769" },

  { id: "batc-A", product: "3.84kWh Batteries Commercial", model: "A", units: 8, size: [1060, 850, 850], unitWeight: 40, color: "#00d4ff" },
  { id: "batc-B", product: "3.84kWh Batteries Commercial", model: "B", units: 6, size: [1060, 850, 670], unitWeight: 40, color: "#007aff" },
  { id: "batc-C", product: "3.84kWh Batteries Commercial", model: "C", units: 4, size: [1060, 530, 850], unitWeight: 40, color: "#6366f1" },

  { id: "batr-A", product: "3.84kWh Batteries Residential", model: "A", units: 16, size: [1350, 1020, 1130], unitWeight: 42, color: "#7d5cff" },
  { id: "batr-B", product: "3.84kWh Batteries Residential", model: "B", units: 12, size: [1350, 1020, 880], unitWeight: 42, color: "#c53cff" },
  { id: "batr-C", product: "3.84kWh Batteries Residential", model: "C", units: 8, size: [1350, 1020, 630], unitWeight: 42, color: "#ff2d55" },

  { id: "bms-A", product: "G4 BMS Residential", model: "A", units: 16, size: [1350, 1020, 1130], unitWeight: 10, color: "#14b8a6" },
  { id: "bms-B", product: "G4 BMS Residential", model: "B", units: 12, size: [1350, 1020, 880], unitWeight: 10, color: "#22d3ee" },
  { id: "bms-C", product: "G4 BMS Residential", model: "C", units: 8, size: [1350, 1020, 630], unitWeight: 10, color: "#2dd4bf" },

  { id: "base-A", product: "G4 Stack Base", model: "A", units: 16, size: [1350, 1020, 770], unitWeight: 15, color: "#fb7185" },
  { id: "base-B", product: "G4 Stack Base", model: "B", units: 12, size: [1350, 1020, 610], unitWeight: 15, color: "#f97316" },
  { id: "base-C", product: "G4 Stack Base", model: "C", units: 8, size: [1020, 680, 770], unitWeight: 15, color: "#f59e0b" },

  { id: "cover-A", product: "AIO Kit Cables Cover", model: "A", units: 16, size: [1350, 1020, 1130], unitWeight: 10, color: "#84cc16" },
  { id: "cover-B", product: "AIO Kit Cables Cover", model: "B", units: 12, size: [1350, 1020, 880], unitWeight: 10, color: "#10b981" },
  { id: "cover-C", product: "AIO Kit Cables Cover", model: "C", units: 8, size: [1350, 1020, 630], unitWeight: 10, color: "#0ea5e9" },

  { id: "cabinet-A", product: "221kWh Cabinet", model: "A", units: 1, size: [1440, 1140, 2475], unitWeight: 2800, color: "#a855f7" },
].map((item) => ({
  ...item,
  sizeM: item.size.map((mm) => mm * MM_TO_M),
  palletWeightKg: PALLET_SELF_WEIGHT_KG + item.unitWeight * item.units,
}));

const PALLET_VISUAL_MAX = PALLET_MODELS.reduce(
  (acc, m) => ({
    l: Math.max(acc.l, m.size[0]),
    w: Math.max(acc.w, m.size[1]),
    h: Math.max(acc.h, m.size[2]),
  }),
  { l: 1, w: 1, h: 1 }
);

const STACK_CONFIG = {
  "inv25-A": { perLayer: 2, layers: 3 },
  "inv25-B": { perLayer: 2, layers: 2 },
  "inv50-A": { perLayer: 1, layers: 3 },
  "inv50-B": { perLayer: 1, layers: 2 },
  "batc-A": { perLayer: 2, layers: 4 },
  "batc-B": { perLayer: 2, layers: 3 },
  "batc-C": { perLayer: 1, layers: 4 },
  "batr-A": { perLayer: 4, layers: 4 },
  "batr-B": { perLayer: 4, layers: 3 },
  "batr-C": { perLayer: 2, layers: 4 },
  "bms-A": { perLayer: 4, layers: 4 },
  "bms-B": { perLayer: 4, layers: 3 },
  "bms-C": { perLayer: 4, layers: 2 },
  "base-A": { perLayer: 4, layers: 4 },
  "base-B": { perLayer: 4, layers: 3 },
  "base-C": { perLayer: 2, layers: 4 },
  "cover-A": { perLayer: 4, layers: 4 },
  "cover-B": { perLayer: 4, layers: 3 },
  "cover-C": { perLayer: 4, layers: 2 },
  "cabinet-A": { perLayer: 1, layers: 1 },
};

const PRODUCT_BUNDLES = [
  {
    id: "res_stack",
    name: "Residential Stack Kit",
    items: [
      { product: "3.84kWh Batteries Residential", units: 1 },
      { product: "G4 BMS Residential", units: 1 },
      { product: "G4 Stack Base", units: 1 },
      { product: "AIO Kit Cables Cover", units: 1 },
    ],
  },
  {
    id: "hybrid_res",
    name: "Hybrid Home Set",
    items: [
      { product: "10-25kW Hybrid Inverter", units: 1 },
      { product: "3.84kWh Batteries Residential", units: 2 },
      { product: "G4 BMS Residential", units: 2 },
      { product: "G4 Stack Base", units: 2 },
      { product: "AIO Kit Cables Cover", units: 2 },
    ],
  },
  {
    id: "hybrid_com",
    name: "Commercial Hybrid Set",
    items: [
      { product: "30-50kW Hybrid Inverter", units: 1 },
      { product: "3.84kWh Batteries Commercial", units: 4 },
    ],
  },
];

const PRODUCT_IMAGE_FILES = {
  "10-25kW Hybrid Inverter": "10-25kw.png",
  "30-50kW Hybrid Inverter": "10-25kw.png",
  "3.84kWh Batteries Commercial": "3-84kwh-batteries-commercial.png",
  "3.84kWh Batteries Residential": "3-84kwh-batteries-residential.png",
  "G4 BMS Residential": "g4-bms-residential.png",
  "G4 Stack Base": "g4-stack-base.png",
  "AIO Kit Cables Cover": "aio-kit-cables-cover.png",
  "221kWh Cabinet": "221kwh-cabinet.png",
};

const refs = {
  containerType: document.getElementById("container-type"),
  containerQuick: document.getElementById("container-quick"),
  maxWeight: document.getElementById("max-weight"),
  planStrategy: document.getElementById("plan-strategy"),
  cscPlateOk: document.getElementById("csc-plate-ok"),
  cscExamOk: document.getElementById("csc-exam-ok"),
  productMix: document.getElementById("product-mix"),
  bundleLibrary: document.getElementById("bundle-library"),
  productsAll: document.getElementById("products-all"),
  productsNone: document.getElementById("products-none"),
  startTour: document.getElementById("start-tour"),
  autoPlan: document.getElementById("auto-plan"),
  palletList: document.getElementById("pallet-list"),
  applyQty: document.getElementById("apply-qty"),
  autoFill: document.getElementById("auto-fill"),
  clearAll: document.getElementById("clear-all"),
  exportPdf: document.getElementById("export-pdf"),
  canvas: document.getElementById("planner-canvas"),
  viewportWrap: document.getElementById("viewport-wrap"),
  statPallets: document.getElementById("stat-pallets"),
  statWeight: document.getElementById("stat-weight"),
  statWeightUtil: document.getElementById("stat-weight-util"),
  statFootprint: document.getElementById("stat-footprint"),
  loadBaseUnits: document.getElementById("load-base-units"),
  loadPalletCount: document.getElementById("load-pallet-count"),
  loadRate: document.getElementById("load-rate"),
  loadRateBar: document.getElementById("load-rate-bar"),
  weightRate: document.getElementById("weight-rate"),
  weightRateBar: document.getElementById("weight-rate-bar"),
  colorLegend: document.getElementById("color-legend"),
  loadProductsList: document.getElementById("load-products-list"),
  algoTrace: document.getElementById("algo-trace"),
  planTotalWrap: document.getElementById("plan-total-wrap"),
  planTotal: document.getElementById("plan-total"),
  containerTabs: document.getElementById("container-tabs"),
  containerData: document.getElementById("container-data"),
  summaryTable: document.getElementById("summary-table"),
  complianceList: document.getElementById("compliance-list"),
  activePlacement: document.getElementById("active-placement"),
};

const state = {
  containerKey: "20GP",
  maxWeightKg: DEFAULT_CONTAINERS["20GP"].maxPayloadKg,
  cscPlateOk: true,
  cscExamOk: true,
  lastComplianceChecks: [],
  enabledProducts: new Set(),
  activeProduct: "",
  productDemand: {},
  bundleQty: {},
  qtyByModel: {},
  selectedModelId: null,
  planStrategy: "max_units",
  initializedDefaults: false,
  containerPlans: [],
  activeContainerPlan: 0,
  isBatchPlanning: false,
  placed: [],
  nextId: 1,
  dropNdc: new THREE.Vector2(0, 0),
};

let scene;
let camera;
let renderer;
let controls;
let containerGroup;
let wireMesh;
let floorMesh;
let usableBoundsMesh;
let raycaster;
const pointerNdc = new THREE.Vector2();
const cargoGroup = new THREE.Group();
let palletPreviewInstances = [];
let viewportResizeObserver = null;
let hintTimer = 0;
let lastContainerSwitchKey = "";
let lastContainerSwitchAt = 0;
let frameTick = 0;
let webglLostAt = 0;
let contextLossRecoveryTimer = 0;
let isRecoveringRenderer = false;
let lastRenderAt = 0;
let sharedEnvMap = null;
const TRACE_MAX_ROWS = 180;
const TOUR_SEEN_KEY = "wattsonic_tour_seen_v1";
let tourActive = false;
let tourStepIndex = 0;
let tourLayer = null;
let tourRing = null;
let tourCard = null;
let tourTitle = null;
let tourBody = null;
let tourStepText = null;
let tourPrevBtn = null;
let tourNextBtn = null;
let tourSkipBtn = null;

const getContainerCfg = () => CONTAINERS[state.containerKey];
const isModelEnabled = (model) => state.enabledProducts.has(model.product);
const ALL_PRODUCTS = Array.from(new Set(PALLET_MODELS.map((m) => m.product)));
const getEnabledProductList = () => ALL_PRODUCTS.filter((p) => state.enabledProducts.has(p));
const getFirstEnabledProduct = () => getEnabledProductList()[0] || "";
const sanitizeDemand = (value) => Math.max(0, Math.floor(Number(value) || 0));
const hasAnyDemand = () => getEnabledProductList().some((product) => sanitizeDemand(state.productDemand[product]) > 0);

const ensureProductDemand = () => {
  ALL_PRODUCTS.forEach((product) => {
    if (!Number.isFinite(state.productDemand[product])) {
      state.productDemand[product] = 0;
    }
  });
};

const ensureBundleQty = () => {
  PRODUCT_BUNDLES.forEach((bundle) => {
    if (!Number.isFinite(state.bundleQty[bundle.id])) {
      state.bundleQty[bundle.id] = 1;
    }
  });
};

const syncContainerQuickUI = () => {
  if (!refs.containerQuick) return;
  refs.containerQuick.querySelectorAll("button").forEach((btn) => {
    btn.classList.toggle("is-active", btn.dataset.containerKey === state.containerKey);
  });
};

const getTourSteps = () => ([
  {
    selector: "#container-type",
    title: "Step 1 · Container Type",
    text: "Choose 20' GP, 40' GP, or 40' HQ. Planning constraints update instantly.",
  },
  {
    selector: "#product-mix",
    title: "Step 2 · Product Demand",
    text: "Enable products and enter demand in pieces for each selected item.",
  },
  {
    selector: "#auto-plan",
    title: "Step 3 · Plan by Demand",
    text: "Click to auto-calculate pallet combinations for maximum loading value.",
  },
  {
    selector: "#pallet-list",
    title: "Step 4 · Pallet Library",
    text: "Review pallet 3D visuals, dimensions, weight, and quantity details.",
  },
  {
    selector: "#viewport-wrap",
    title: "Step 5 · Container 3D",
    text: "Inspect the loading layout in 3D and verify occupancy and stacking status.",
  },
  {
    selector: "#export-pdf",
    title: "Step 6 · Download PDF",
    text: "Export a clean customer-ready loading plan with product, pallet, and 3D snapshot.",
  },
]);

const ensureTourUi = () => {
  if (tourLayer) return;
  tourLayer = document.createElement("div");
  tourLayer.className = "tour-layer";
  tourLayer.innerHTML = `
    <div class="tour-ring"></div>
    <div class="tour-card">
      <div class="tour-step"></div>
      <h3 class="tour-title"></h3>
      <p class="tour-body"></p>
      <div class="tour-actions">
        <button type="button" class="tour-btn ghost" data-tour-action="prev">Back</button>
        <button type="button" class="tour-btn ghost" data-tour-action="skip">Skip</button>
        <button type="button" class="tour-btn primary" data-tour-action="next">Next</button>
      </div>
    </div>
  `;
  document.body.appendChild(tourLayer);
  tourRing = tourLayer.querySelector(".tour-ring");
  tourCard = tourLayer.querySelector(".tour-card");
  tourTitle = tourLayer.querySelector(".tour-title");
  tourBody = tourLayer.querySelector(".tour-body");
  tourStepText = tourLayer.querySelector(".tour-step");
  tourPrevBtn = tourLayer.querySelector('[data-tour-action="prev"]');
  tourNextBtn = tourLayer.querySelector('[data-tour-action="next"]');
  tourSkipBtn = tourLayer.querySelector('[data-tour-action="skip"]');
  tourPrevBtn?.addEventListener("click", () => moveTourStep(-1));
  tourNextBtn?.addEventListener("click", () => moveTourStep(1));
  tourSkipBtn?.addEventListener("click", () => closeTour(true));
};

const getVisibleTourTarget = (selector) => {
  const el = document.querySelector(selector);
  if (!el) return null;
  const rect = el.getBoundingClientRect();
  if (rect.width <= 0 || rect.height <= 0) return null;
  return el;
};

const positionTour = () => {
  if (!tourActive) return;
  const steps = getTourSteps();
  const step = steps[tourStepIndex];
  if (!step) return;
  const target = getVisibleTourTarget(step.selector);
  if (!target) return;

  target.scrollIntoView({ behavior: "smooth", block: "center", inline: "nearest" });
  const rect = target.getBoundingClientRect();
  const pad = 8;
  tourRing.style.left = `${Math.max(8, rect.left - pad)}px`;
  tourRing.style.top = `${Math.max(8, rect.top - pad)}px`;
  tourRing.style.width = `${Math.max(32, rect.width + pad * 2)}px`;
  tourRing.style.height = `${Math.max(32, rect.height + pad * 2)}px`;

  const cardW = 360;
  const viewportW = window.innerWidth;
  const viewportH = window.innerHeight;
  let cardX = Math.min(viewportW - cardW - 16, Math.max(16, rect.left));
  let cardY = rect.bottom + 14;
  if (cardY + 170 > viewportH - 12) {
    cardY = Math.max(12, rect.top - 182);
  }
  if (rect.width > viewportW * 0.72) {
    cardX = Math.max(16, (viewportW - cardW) / 2);
  }
  tourCard.style.left = `${cardX}px`;
  tourCard.style.top = `${cardY}px`;
};

const renderTourStep = () => {
  if (!tourActive) return;
  const steps = getTourSteps();
  const step = steps[tourStepIndex];
  if (!step) return;
  tourStepText.textContent = `${tourStepIndex + 1} / ${steps.length}`;
  tourTitle.textContent = step.title;
  tourBody.textContent = step.text;
  tourPrevBtn.disabled = tourStepIndex === 0;
  tourNextBtn.textContent = tourStepIndex === steps.length - 1 ? "Done" : "Next";
  positionTour();
};

const openTour = (force = false) => {
  if (!force && window.localStorage.getItem(TOUR_SEEN_KEY) === "1") return;
  ensureTourUi();
  tourActive = true;
  tourStepIndex = 0;
  tourLayer.classList.add("is-active");
  renderTourStep();
};

const closeTour = (markSeen = false) => {
  if (!tourLayer) return;
  tourActive = false;
  tourLayer.classList.remove("is-active");
  if (markSeen) {
    window.localStorage.setItem(TOUR_SEEN_KEY, "1");
  }
};

const moveTourStep = (delta) => {
  const steps = getTourSteps();
  const next = tourStepIndex + delta;
  if (next < 0) return;
  if (next >= steps.length) {
    closeTour(true);
    return;
  }
  tourStepIndex = next;
  renderTourStep();
};

const showPlacementHint = (message, kind = "info") => {
  if (!refs.activePlacement) return;
  refs.activePlacement.textContent = message;
  refs.activePlacement.dataset.kind = kind;
  if (hintTimer) clearTimeout(hintTimer);
  hintTimer = window.setTimeout(() => {
    const current = state.selectedModelId ? PALLET_MODELS.find((m) => m.id === state.selectedModelId) : null;
    refs.activePlacement.textContent = `Current Placement: ${current ? `${current.product} ${current.model}` : "None"}`;
    refs.activePlacement.dataset.kind = "";
  }, 1800);
};

const clearAlgoTrace = () => {
  if (!refs.algoTrace) return;
  refs.algoTrace.innerHTML = "";
};

const logAlgoTrace = (message, kind = "info") => {
  if (!refs.algoTrace) return;
  const row = document.createElement("div");
  row.className = `algo-row ${kind}`;
  row.textContent = message;
  refs.algoTrace.prepend(row);
  while (refs.algoTrace.children.length > TRACE_MAX_ROWS) {
    refs.algoTrace.removeChild(refs.algoTrace.lastChild);
  }
};

const ensureContainerVisible = () => {
  if (!scene) return;
  const presentInScene = !!containerGroup && scene.children.includes(containerGroup);
  const hasGeometry = !!containerGroup && containerGroup.children.length > 0;
  if (!presentInScene || !hasGeometry) {
    rebuildContainer();
  }
};

const resetMainView = () => {
  if (!camera || !controls) return;
  camera.position.set(9, 6, 8);
  const bounds = getContainerBounds();
  controls.target.set(0, bounds.maxY * 0.4, 0);
  controls.update();
};

const syncViewportAndScene = () => {
  if (!renderer || !camera || !scene) return;
  onResize();
  ensureContainerVisible();
  renderer.render(scene, camera);
};

const queueViewportRefresh = (passes = 3) => {
  let remaining = Math.max(1, passes);
  const step = () => {
    syncViewportAndScene();
    remaining -= 1;
    if (remaining > 0) requestAnimationFrame(step);
  };
  requestAnimationFrame(step);
};

const configureMainRenderer = (instance) => {
  instance.setPixelRatio(Math.min(window.devicePixelRatio, 2));
  instance.physicallyCorrectLights = true;
  instance.toneMapping = THREE.ACESFilmicToneMapping;
  instance.toneMappingExposure = 1.08;
  instance.outputColorSpace = THREE.SRGBColorSpace;
};

const recoverMainRenderer = (reason = "unknown") => {
  if (!scene || !camera || !refs.canvas || isRecoveringRenderer) return;
  isRecoveringRenderer = true;
  try {
    const prev = renderer;
    renderer = new THREE.WebGLRenderer({ canvas: refs.canvas, antialias: true, alpha: false, preserveDrawingBuffer: true });
    configureMainRenderer(renderer);
    if (!sharedEnvMap) {
      const pmrem = new THREE.PMREMGenerator(renderer);
      sharedEnvMap = pmrem.fromScene(new RoomEnvironment(), 0.04).texture;
      pmrem.dispose();
    }
    scene.environment = sharedEnvMap;
    prev?.dispose?.();
    onResize();
    rebuildContainer();
    updateStats();
    queueViewportRefresh(3);
    showPlacementHint(`3D renderer recovered (${reason}).`, "ok");
  } catch (err) {
    console.error("Renderer recovery failed:", err);
  } finally {
    isRecoveringRenderer = false;
  }
};

const invalidateContainerPlans = () => {
  if (state.isBatchPlanning) return;
  state.containerPlans = [];
  state.activeContainerPlan = 0;
  renderContainerTabs();
};

const snapshotCurrentPlacements = () =>
  state.placed.map((p) => ({
    modelId: p.modelId,
    x: p.x,
    y: p.y,
    z: p.z,
    len: p.len,
    wid: p.wid,
    hei: p.hei,
    rotationY: p.rotationY || 0,
    weightKg: p.weightKg,
    product: p.product,
    model: p.model,
  }));

const clearActivePlacements = (withStats = true) => {
  const dispose = [...state.placed];
  dispose.forEach((p) => {
    cargoGroup.remove(p.mesh);
    disposeObject3D(p.mesh);
  });
  state.placed = [];
  if (withStats) updateStats();
};

const restorePlacementSnapshot = (snap) => {
  const model = PALLET_MODELS.find((m) => m.id === snap.modelId);
  if (!model) return;
  const geom = new THREE.BoxGeometry(snap.len, snap.hei, snap.wid);
  const mat = new THREE.MeshStandardMaterial({ color: model.color, roughness: 0.3, metalness: 0.05 });
  const mesh = new THREE.Mesh(geom, mat);
  mesh.position.set(snap.x + snap.len / 2, snap.y + snap.hei / 2, snap.z + snap.wid / 2);
  mesh.rotation.y = snap.rotationY || 0;
  const edge = new THREE.LineSegments(new THREE.EdgesGeometry(geom), new THREE.LineBasicMaterial({ color: "#1f2635" }));
  mesh.add(edge);
  const id = state.nextId++;
  mesh.userData.placementId = id;
  cargoGroup.add(mesh);
  state.placed.push({
    id,
    modelId: snap.modelId,
    product: snap.product,
    model: snap.model,
    weightKg: snap.weightKg,
    x: snap.x,
    y: snap.y,
    z: snap.z,
    len: snap.len,
    wid: snap.wid,
    hei: snap.hei,
    rotationY: snap.rotationY || 0,
    mesh,
  });
};

const loadContainerPlan = (index) => {
  if (!state.containerPlans.length) return;
  const clamped = Math.max(0, Math.min(index, state.containerPlans.length - 1));
  state.activeContainerPlan = clamped;
  clearActivePlacements(false);
  state.containerPlans[clamped].placements.forEach((snap) => restorePlacementSnapshot(snap));
  renderContainerTabs();
  updateStats();
  resetMainView();
  queueViewportRefresh(2);
};

const renderContainerTabs = () => {
  if (!refs.containerTabs || !refs.planTotal || !refs.planTotalWrap) return;
  const total = state.containerPlans.length;
  refs.planTotalWrap.hidden = total <= 1;
  refs.planTotal.textContent = String(Math.max(1, total || 1));
  refs.containerTabs.innerHTML = "";
  if (total <= 1) return;
  state.containerPlans.forEach((plan, idx) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = `container-tab ${idx === state.activeContainerPlan ? "is-active" : ""}`;
    btn.textContent = `Container ${idx + 1}`;
    btn.title = `${plan.baseUnits} units · ${plan.pallets} pallets`;
    btn.addEventListener("click", () => loadContainerPlan(idx));
    refs.containerTabs.appendChild(btn);
  });
};

const productColor = (product) => PALLET_MODELS.find((m) => m.product === product)?.color || "#90a2bf";
const productImageUrl = (product) => {
  const file = PRODUCT_IMAGE_FILES[product];
  return file ? `./assets/products/${file}` : "";
};

const disposeObject3D = (root) => {
  if (!root) return;
  root.traverse((node) => {
    if (node.geometry?.dispose) node.geometry.dispose();
    if (node.material) {
      if (Array.isArray(node.material)) {
        node.material.forEach((m) => m?.dispose?.());
      } else {
        node.material.dispose?.();
      }
    }
  });
};

const applyContainerType = (nextKey) => {
  if (!CONTAINERS[nextKey]) return;
  const now = Date.now();
  if (nextKey === lastContainerSwitchKey && now - lastContainerSwitchAt < 260) return;
  lastContainerSwitchKey = nextKey;
  lastContainerSwitchAt = now;

  if (nextKey === state.containerKey) {
    queueViewportRefresh(2);
    return;
  }

  const hadDemand = hasAnyDemand();
  state.containerKey = nextKey;
  const cfg = getContainerCfg();
  state.maxWeightKg = cfg.maxPayloadKg;
  refs.containerType.value = nextKey;
  refs.maxWeight.value = String(cfg.maxPayloadKg);
  clearAllCargo();
  rebuildContainer();
  resetMainView();
  if (renderer?.getContext?.()?.isContextLost?.()) {
    recoverMainRenderer("container-switch");
  }
  queueViewportRefresh(3);
  renderContainerData();
  syncContainerQuickUI();
  if (hadDemand) {
    logAlgoTrace(`Container switched to ${cfg.label}. Re-running demand plan.`, "info");
    const ok = autoPlanByProductDemand();
    queueViewportRefresh(3);
    resetMainView();
    if (ok) {
      showPlacementHint(`Switched to ${cfg.label}. Plan recalculated by demand.`, "ok");
    } else {
      showPlacementHint(`Switched to ${cfg.label}. Re-plan failed under current constraints.`, "warn");
    }
    return;
  }
  queueViewportRefresh(2);
  resetMainView();
  renderPalletCards();
};

const loadContainerCatalog = async () => {
  try {
    const res = await fetch("./container-models.json", { cache: "no-store" });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    if (!data || typeof data !== "object") throw new Error("Invalid catalog");
    CONTAINERS = data;
  } catch (err) {
    CONTAINERS = { ...DEFAULT_CONTAINERS };
    console.warn("Failed to load container-models.json, fallback to defaults.", err);
  }
  if (!CONTAINERS[state.containerKey]) {
    state.containerKey = Object.keys(CONTAINERS)[0];
  }
  state.maxWeightKg = getContainerCfg().maxPayloadKg;
};

const setupUI = () => {
  refs.containerType.innerHTML = "";
  Object.entries(CONTAINERS).forEach(([key, cfg]) => {
    const opt = document.createElement("option");
    opt.value = key;
    opt.textContent = `${cfg.label} (${cfg.lengthMm}×${cfg.widthMm}×${cfg.heightMm} mm)`;
    refs.containerType.appendChild(opt);
  });
  refs.containerQuick.innerHTML = "";
  Object.entries(CONTAINERS).forEach(([key, cfg]) => {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.dataset.containerKey = key;
    btn.textContent = cfg.label;
    btn.addEventListener("click", () => applyContainerType(key));
    refs.containerQuick.appendChild(btn);
  });
  if (!CONTAINERS[state.containerKey]) {
    state.containerKey = Object.keys(CONTAINERS)[0];
  }
  if (!state.initializedDefaults) {
    state.enabledProducts.clear();
    state.activeProduct = "";
    state.initializedDefaults = true;
  }
  ensureProductDemand();
  ensureBundleQty();
  if (!state.activeProduct || !state.enabledProducts.has(state.activeProduct)) {
    state.activeProduct = getFirstEnabledProduct();
  }
  refs.containerType.value = state.containerKey;
  refs.maxWeight.value = String(state.maxWeightKg);
  if (refs.planStrategy) refs.planStrategy.value = state.planStrategy;
  renderContainerData();
  syncContainerQuickUI();
  renderProductMix();
  renderBundleLibrary();
  renderPalletCards();
  renderContainerTabs();

  refs.containerType.addEventListener("change", () => applyContainerType(refs.containerType.value));

  refs.planStrategy?.addEventListener("change", () => {
    state.planStrategy = refs.planStrategy.value || "max_units";
  });

  refs.startTour?.addEventListener("click", () => openTour(true));
  refs.applyQty?.addEventListener("click", fillByQuantity);
  refs.autoFill?.addEventListener("click", autoFillGreedy);
  refs.autoPlan?.addEventListener("click", autoPlanByProductDemand);
  refs.exportPdf?.addEventListener("click", exportLoadingPlanPdfSafe);
  refs.clearAll.addEventListener("click", clearAllCargo);
  refs.cscPlateOk?.addEventListener("change", () => {
    state.cscPlateOk = refs.cscPlateOk.checked;
    updateStats();
  });
  refs.cscExamOk?.addEventListener("change", () => {
    state.cscExamOk = refs.cscExamOk.checked;
    updateStats();
  });
  refs.productsAll?.addEventListener("click", () => {
    state.enabledProducts = new Set(ALL_PRODUCTS);
    ensureProductDemand();
    renderProductMix();
    renderBundleLibrary();
    state.activeProduct = state.activeProduct && state.enabledProducts.has(state.activeProduct)
      ? state.activeProduct
      : getFirstEnabledProduct();
    syncPalletAvailabilityUI();
    renderPalletCards();
  });
  refs.productsNone?.addEventListener("click", () => {
    state.enabledProducts.clear();
    state.activeProduct = "";
    state.selectedModelId = null;
    ALL_PRODUCTS.forEach((product) => {
      state.productDemand[product] = 0;
    });
    if (refs.activePlacement) refs.activePlacement.textContent = "Current Placement: None";
    renderProductMix();
    renderBundleLibrary();
    syncPalletAvailabilityUI();
    renderPalletCards();
  });

  refs.viewportWrap.addEventListener("dragover", (event) => {
    event.preventDefault();
    updateDropNdcFromPointer(event.clientX, event.clientY, state.dropNdc);
  });

  refs.viewportWrap.addEventListener("drop", (event) => {
    event.preventDefault();
    updateDropNdcFromPointer(event.clientX, event.clientY, state.dropNdc);
    const plain = event.dataTransfer?.getData("text/plain") || "";
    const bundlePayload = event.dataTransfer?.getData("application/x-wattsonic-bundle")
      || (plain.startsWith("bundle:") ? plain.slice(7) : "");
    if (bundlePayload) {
      const [bundleId, qtyText] = bundlePayload.split("|");
      const qty = Math.max(1, sanitizeDemand(qtyText) || 1);
      const ok = applyBundleAndReplan(bundleId, qty);
      if (!ok) showPlacementHint("Bundle planning failed under current constraints.", "warn");
      return;
    }
    const modelId = event.dataTransfer?.getData("application/x-wattsonic-model")
      || (plain.startsWith("model:") ? plain.slice(6) : plain);
    const product = event.dataTransfer?.getData("application/x-wattsonic-product")
      || (plain.startsWith("product:") ? plain.slice(8) : "");
    if (product) {
      const ok = placeBestPalletForProduct(product, state.dropNdc);
      if (!ok) showPlacementHint(`Cannot place ${product}: no valid pallet fits now.`, "warn");
      return;
    }
    const model = PALLET_MODELS.find((item) => item.id === modelId);
    if (!model || !isModelEnabled(model)) return;
    let ok = dropModelAtPointer(model, state.dropNdc);
    if (!ok) ok = placeModelSmart(model);
    if (ok) {
      showPlacementHint(
        `Loaded: ${model.product} ${model.model} · ${state.placed.length} pallets · ${Math.round(currentWeightKg())} kg`,
        "ok"
      );
    } else {
      showPlacementHint(`Cannot place: no space or payload left for ${model.product} ${model.model}`, "warn");
    }
  });

  refs.viewportWrap.addEventListener("pointerdown", onViewportPointerDown);
};

const setSelectedModel = (modelId) => {
  state.selectedModelId = modelId;
  const model = PALLET_MODELS.find((item) => item.id === modelId);
  if (refs.activePlacement) {
    refs.activePlacement.textContent = `Current Placement: ${model ? `${model.product} ${model.model}` : "None"}`;
  }
  refs.palletList.querySelectorAll(".pallet-item").forEach((el) => {
    el.classList.toggle("is-selected", el.dataset.modelId === modelId);
  });
};

const renderContainerData = () => {
  const cfg = getContainerCfg();
  if (!refs.containerData || !cfg) return;
  refs.containerData.innerHTML = `
    <div class="container-data-row"><span>Type</span><strong>${cfg.label}</strong></div>
    <div class="container-data-row"><span>Internal Size</span><strong>${cfg.lengthMm} × ${cfg.widthMm} × ${cfg.heightMm} mm</strong></div>
    <div class="container-data-row"><span>Tare Weight</span><strong>${cfg.tareKg} kg</strong></div>
    <div class="container-data-row"><span>Max Payload</span><strong>${cfg.maxPayloadKg} kg</strong></div>
    <div class="container-data-row"><span>Max Gross</span><strong>${cfg.maxGrossKg} kg</strong></div>
  `;
};

const renderProductMix = () => {
  if (!refs.productMix) return;
  ensureProductDemand();
  const products = ALL_PRODUCTS;
  refs.productMix.innerHTML = "";
  products.forEach((product) => {
    const row = document.createElement("div");
    row.className = "product-mix-item";
    row.draggable = true;
    if (product === state.activeProduct) row.classList.add("is-active");
    const checked = state.enabledProducts.has(product);
    const demand = sanitizeDemand(state.productDemand[product]);
    const imageUrl = productImageUrl(product);
    const shortTag = product
      .split(" ")
      .filter(Boolean)
      .slice(0, 2)
      .map((part) => part[0]?.toUpperCase() || "")
      .join("");
    row.innerHTML = `
      <div class="product-card-image-wrap">
        <img class="product-thumb" src="${imageUrl}" alt="${product}" loading="lazy" />
        <div class="product-thumb-fallback">${shortTag}</div>
      </div>
      <div class="product-card-content">
        <label class="product-toggle">
          <input type="checkbox" data-product="${product}" ${checked ? "checked" : ""} />
          <span>${product}</span>
        </label>
        <label class="product-demand">
          <span>Demand</span>
          <input type="number" min="0" step="1" value="${demand}" data-demand-product="${product}" />
          <em>pieces</em>
        </label>
      </div>
    `;
    const thumb = row.querySelector(".product-thumb");
    if (!imageUrl) {
      row.classList.add("no-image");
    }
    thumb?.addEventListener("error", () => {
      row.classList.add("no-image");
    });
    thumb?.addEventListener("load", () => {
      row.classList.remove("no-image");
    });
    const box = row.querySelector(`[data-product="${product}"]`);
    row.addEventListener("dragstart", (event) => {
      event.dataTransfer?.setData("application/x-wattsonic-product", product);
      event.dataTransfer?.setData("text/plain", `product:${product}`);
    });
    box?.addEventListener("click", (e) => e.stopPropagation());
    box?.addEventListener("dragstart", (e) => e.preventDefault());
    box?.addEventListener("change", () => {
      if (box.checked) {
        state.enabledProducts.add(product);
        if (!state.activeProduct || !state.enabledProducts.has(state.activeProduct)) {
          state.activeProduct = product;
        }
      } else {
        state.enabledProducts.delete(product);
        if (state.activeProduct === product) {
          state.activeProduct = getFirstEnabledProduct();
        }
      }
      renderProductMix();
      syncPalletAvailabilityUI();
      renderPalletCards();
    });
    const demandInput = row.querySelector(`[data-demand-product="${product}"]`);
    demandInput?.addEventListener("click", (e) => e.stopPropagation());
    demandInput?.addEventListener("dragstart", (e) => e.preventDefault());
    demandInput?.addEventListener("change", () => {
      state.productDemand[product] = sanitizeDemand(demandInput.value);
      demandInput.value = String(state.productDemand[product]);
    });
    row.addEventListener("click", () => {
      if (!state.enabledProducts.has(product)) return;
      state.activeProduct = product;
      renderProductMix();
      renderPalletCards();
    });
    refs.productMix.appendChild(row);
  });
};

const renderBundleLibrary = () => {
  if (!refs.bundleLibrary) return;
  ensureBundleQty();
  refs.bundleLibrary.innerHTML = "";
  PRODUCT_BUNDLES.forEach((bundle) => {
    const qty = Math.max(1, sanitizeDemand(state.bundleQty[bundle.id]) || 1);
    const card = document.createElement("article");
    card.className = "bundle-card";
    card.draggable = true;
    card.dataset.bundleId = bundle.id;
    const detail = bundle.items.map((item) => `${item.product} × ${item.units}`).join(" · ");
    card.innerHTML = `
      <div class="bundle-title">${bundle.name}</div>
      <div class="bundle-items">${detail}</div>
      <label class="bundle-qty">
        <span>Pack Qty</span>
        <input type="number" min="1" step="1" value="${qty}" data-bundle-qty="${bundle.id}" />
      </label>
    `;
    const qtyInput = card.querySelector(`[data-bundle-qty="${bundle.id}"]`);
    qtyInput?.addEventListener("click", (e) => e.stopPropagation());
    qtyInput?.addEventListener("change", () => {
      state.bundleQty[bundle.id] = Math.max(1, sanitizeDemand(qtyInput.value) || 1);
      qtyInput.value = String(state.bundleQty[bundle.id]);
    });
    card.addEventListener("dragstart", (event) => {
      const currentQty = Math.max(1, sanitizeDemand(state.bundleQty[bundle.id]) || 1);
      const payload = `${bundle.id}|${currentQty}`;
      event.dataTransfer?.setData("application/x-wattsonic-bundle", payload);
      event.dataTransfer?.setData("text/plain", `bundle:${payload}`);
    });
    card.addEventListener("click", () => {
      const currentQty = Math.max(1, sanitizeDemand(state.bundleQty[bundle.id]) || 1);
      const ok = applyBundleAndReplan(bundle.id, currentQty);
      if (!ok) showPlacementHint("Bundle planning failed under current constraints.", "warn");
    });
    refs.bundleLibrary.appendChild(card);
  });
};

const syncPalletAvailabilityUI = () => {
  if (state.selectedModelId) {
    const current = PALLET_MODELS.find((m) => m.id === state.selectedModelId);
    if (current && !isModelEnabled(current)) {
      state.selectedModelId = null;
      if (refs.activePlacement) refs.activePlacement.textContent = "Current Placement: None";
    }
  }
};

const disposePalletPreviews = () => {
  palletPreviewInstances.forEach((inst) => inst.dispose());
  palletPreviewInstances = [];
};

const createPalletPreview = (canvas, model) => {
  const rendererLocal = new THREE.WebGLRenderer({ canvas, antialias: true, alpha: true });
  rendererLocal.setPixelRatio(Math.min(window.devicePixelRatio, 2));

  const sceneLocal = new THREE.Scene();
  const cameraLocal = new THREE.PerspectiveCamera(36, 1, 0.01, 20);
  cameraLocal.position.set(0, 0.92, 2.45);
  cameraLocal.lookAt(0, 0.5, 0);

  const key = new THREE.HemisphereLight("#ffffff", "#dae6ff", 1.0);
  const fill = new THREE.DirectionalLight("#ffffff", 0.72);
  fill.position.set(2, 3, 1.5);
  sceneLocal.add(key, fill);

  const group = new THREE.Group();
  sceneLocal.add(group);

  const stack = STACK_CONFIG[model.id] || { perLayer: Math.max(1, Math.floor(model.units / 2)), layers: 2 };
  const patternByPerLayer = {
    1: { cols: 1, rows: 1 },
    2: { cols: 2, rows: 1 },
    3: { cols: 3, rows: 1 },
    4: { cols: 2, rows: 2 },
  };
  const layout = patternByPerLayer[stack.perLayer] || { cols: Math.ceil(Math.sqrt(stack.perLayer)), rows: Math.ceil(stack.perLayer / Math.ceil(Math.sqrt(stack.perLayer))) };
  const cols = layout.cols;
  const rows = layout.rows;

  const modelL = model.size[0] / 1000;
  const modelW = model.size[1] / 1000;
  const modelH = model.size[2] / 1000;

  // Keep geometric ratio true to mm data while fitting card viewport.
  const footprintScale = 0.74;
  const baseL = modelL * footprintScale;
  const baseW = modelW * footprintScale;
  const totalH = modelH * footprintScale;

  const palletDeckH = 0.038;
  const palletFootH = 0.075;
  const cargoStartY = palletDeckH + palletFootH + 0.01;
  const layerH = Math.max(0.06, totalH / stack.layers);

  const marginL = baseL * 0.045;
  const marginW = baseW * 0.045;
  const gapL = baseL * 0.02;
  const gapW = baseW * 0.02;
  const itemL = (baseL - marginL * 2 - gapL * (cols - 1)) / cols;
  const itemW = (baseW - marginW * 2 - gapW * (rows - 1)) / rows;

  const cartonGeom = new THREE.BoxGeometry(itemL, layerH * 0.94, itemW);
  const cartonMat = new THREE.MeshStandardMaterial({ color: "#e5d2af", roughness: 0.78, metalness: 0.02 });
  const edgeMat = new THREE.LineBasicMaterial({ color: "#b89c6d" });

  for (let l = 0; l < stack.layers; l += 1) {
    for (let i = 0; i < stack.perLayer; i += 1) {
      const c = i % cols;
      const r = Math.floor(i / cols);
      const mesh = new THREE.Mesh(cartonGeom, cartonMat);
      mesh.position.set(
        -baseL / 2 + marginL + c * (itemL + gapL) + itemL / 2,
        cargoStartY + (l + 0.5) * layerH,
        -baseW / 2 + marginW + r * (itemW + gapW) + itemW / 2
      );
      const edges = new THREE.LineSegments(new THREE.EdgesGeometry(cartonGeom), edgeMat);
      mesh.add(edges);
      group.add(mesh);
    }
  }

  const deckGeom = new THREE.BoxGeometry(baseL + 0.045, palletDeckH, baseW + 0.045);
  const deckMat = new THREE.MeshStandardMaterial({ color: "#d2b789", roughness: 0.82, metalness: 0.01 });
  const deck = new THREE.Mesh(deckGeom, deckMat);
  deck.position.y = palletFootH + palletDeckH / 2;
  group.add(deck);

  const runnerGeom = new THREE.BoxGeometry(baseL * 0.27, palletFootH, baseW * 0.14);
  const runnerMat = new THREE.MeshStandardMaterial({ color: "#c9aa75", roughness: 0.84, metalness: 0.01 });
  const footPos = [
    [-baseL * 0.33, palletFootH / 2, -baseW * 0.30],
    [0, palletFootH / 2, -baseW * 0.30],
    [baseL * 0.33, palletFootH / 2, -baseW * 0.30],
    [-baseL * 0.33, palletFootH / 2, baseW * 0.30],
    [0, palletFootH / 2, baseW * 0.30],
    [baseL * 0.33, palletFootH / 2, baseW * 0.30],
  ];
  footPos.forEach(([x, y, z]) => {
    const foot = new THREE.Mesh(runnerGeom, runnerMat);
    foot.position.set(x, y, z);
    group.add(foot);
  });

  group.rotation.x = 0;
  group.rotation.y = 0;

  let isDragging = false;
  let lx = 0;
  let ly = 0;
  let raf = 0;
  let cw = 0;
  let ch = 0;

  const onDown = (e) => {
    isDragging = true;
    lx = e.clientX;
    ly = e.clientY;
    canvas.style.cursor = "grabbing";
  };
  const onMove = (e) => {
    if (!isDragging) return;
    const dx = e.clientX - lx;
    lx = e.clientX;
    ly = e.clientY;
    group.rotation.y += dx * 0.0045;
  };
  const onUp = () => {
    isDragging = false;
    canvas.style.cursor = "grab";
  };

  canvas.addEventListener("pointerdown", onDown);
  window.addEventListener("pointermove", onMove);
  window.addEventListener("pointerup", onUp);
  canvas.style.cursor = "grab";

  const tick = () => {
    const w = Math.max(120, canvas.clientWidth | 0);
    const h = Math.max(90, canvas.clientHeight | 0);
    if (w !== cw || h !== ch) {
      cw = w;
      ch = h;
      rendererLocal.setSize(w, h, false);
      cameraLocal.aspect = w / h;
      cameraLocal.updateProjectionMatrix();
    }
    rendererLocal.render(sceneLocal, cameraLocal);
    raf = requestAnimationFrame(tick);
  };
  tick();

  return {
    dispose() {
      cancelAnimationFrame(raf);
      canvas.removeEventListener("pointerdown", onDown);
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      group.traverse((obj) => {
        if (obj.isMesh && obj.geometry) obj.geometry.dispose();
        if (obj.isMesh && obj.material) {
          if (Array.isArray(obj.material)) obj.material.forEach((m) => m.dispose?.());
          else obj.material.dispose?.();
        }
      });
      rendererLocal.forceContextLoss?.();
      rendererLocal.dispose();
      canvas.width = 1;
      canvas.height = 1;
    },
  };
};

const renderPalletCards = () => {
  disposePalletPreviews();
  refs.palletList.innerHTML = "";
  if (!state.activeProduct && state.enabledProducts.size > 0) {
    state.activeProduct = getFirstEnabledProduct();
    renderProductMix();
  }
  if (!state.activeProduct || !state.enabledProducts.has(state.activeProduct)) {
    refs.palletList.innerHTML = `<div class="pallet-empty">Select a product from Product Mix to view pallet specs.</div>`;
    return;
  }

  const cfg = getContainerCfg();
  const floorAreaMm2 = cfg.lengthMm * cfg.widthMm;
  const models = PALLET_MODELS.filter((m) => m.product === state.activeProduct && isModelEnabled(m));

  models.forEach((model) => {
    const theoreticalByArea = Math.max(0, Math.floor(floorAreaMm2 / (model.size[0] * model.size[1])));
    const theoreticalByWeight = Math.max(0, Math.floor(state.maxWeightKg / model.palletWeightKg));
    const theoreticalMax = Math.min(theoreticalByArea, theoreticalByWeight);

    const card = document.createElement("article");
    card.className = "pallet-item pallet-spec-card";
    card.draggable = false;
    card.dataset.modelId = model.id;
    if (state.selectedModelId === model.id) card.classList.add("is-selected");

    card.innerHTML = `
      <div class="pallet-spec-header">
        <div class="pallet-spec-title">
          <span class="swatch" style="background:${model.color}"></span>
          <strong>Pallet ${model.model}</strong>
        </div>
        <div class="drag-handle" draggable="true" data-drag-model="${model.id}" title="Drag into container">Drag</div>
      </div>
      <div class="pallet-spec-body">
        <div class="pallet-spec-info">
          <div class="spec-line"><span>Included Product</span><strong>${model.product}</strong></div>
          <div class="spec-line"><span>Units per Pallet</span><strong>${model.units} pcs</strong></div>
          <div class="spec-line"><span>Stacking Method</span><strong>${STACK_CONFIG[model.id]?.perLayer || 1} pcs/layer × ${STACK_CONFIG[model.id]?.layers || 1} layers</strong></div>
          <div class="spec-line"><span>Net Pallet Weight</span><strong>${PALLET_SELF_WEIGHT_KG}+${model.unitWeight}×${model.units}=${Math.round(model.palletWeightKg)} kg</strong></div>
          <div class="spec-line"><span>Pallet Size (L×W×H)</span><strong>${model.size[0]} × ${model.size[1]} × ${model.size[2]} mm</strong></div>
          <div class="spec-line"><span>Theoretical Max / ${cfg.label}</span><strong>${theoreticalMax} pallets (single-layer)</strong></div>
          <label class="spec-qty">Quantity
            <input class="qty-input" type="number" min="0" step="1" value="${state.qtyByModel[model.id] ?? 0}" data-qty-model="${model.id}" />
          </label>
        </div>
        <div class="pallet-visual">
          <canvas class="pallet-canvas" data-pallet-canvas="${model.id}"></canvas>
          <div class="pallet-canvas-meta">
            <span>${model.size[1]} mm</span>
            <span>${model.size[0]} mm</span>
            <strong>${Math.round(model.palletWeightKg)} kg</strong>
          </div>
          <div class="pallet-canvas-tip">Drag to rotate</div>
        </div>
      </div>
    `;

    card.addEventListener("click", () => setSelectedModel(model.id));

    const dragHandle = card.querySelector(`[data-drag-model="${model.id}"]`);
    dragHandle?.addEventListener("dragstart", (event) => {
      event.dataTransfer?.setData("application/x-wattsonic-model", model.id);
      event.dataTransfer?.setData("text/plain", model.id);
      setSelectedModel(model.id);
    });

    const qtyInput = card.querySelector(`[data-qty-model="${model.id}"]`);
    qtyInput?.addEventListener("click", (e) => e.stopPropagation());
    qtyInput?.addEventListener("change", () => {
      state.qtyByModel[model.id] = Math.max(0, Number(qtyInput.value) || 0);
    });

    refs.palletList.appendChild(card);
    const previewCanvas = card.querySelector(`[data-pallet-canvas="${model.id}"]`);
    if (previewCanvas) {
      previewCanvas.addEventListener("click", (e) => e.stopPropagation());
      previewCanvas.addEventListener("dragstart", (e) => e.preventDefault());
      palletPreviewInstances.push(createPalletPreview(previewCanvas, model));
    }
  });
};

const init3D = () => {
  scene = new THREE.Scene();
  scene.background = new THREE.Color("#edf3ff");

  camera = new THREE.PerspectiveCamera(52, 1, 0.1, 150);
  camera.position.set(9, 6, 8);

  renderer = new THREE.WebGLRenderer({ canvas: refs.canvas, antialias: true, alpha: false, preserveDrawingBuffer: true });
  configureMainRenderer(renderer);
  const pmrem = new THREE.PMREMGenerator(renderer);
  sharedEnvMap = pmrem.fromScene(new RoomEnvironment(), 0.04).texture;
  scene.environment = sharedEnvMap;
  pmrem.dispose();
  refs.canvas.addEventListener("webglcontextlost", (event) => {
    event.preventDefault();
    webglLostAt = performance.now();
    showPlacementHint("3D context lost. Recovering...", "warn");
    if (contextLossRecoveryTimer) clearTimeout(contextLossRecoveryTimer);
    contextLossRecoveryTimer = window.setTimeout(() => {
      const lost = renderer?.getContext?.()?.isContextLost?.() || false;
      if (lost) recoverMainRenderer("context-lost");
    }, 260);
  });
  refs.canvas.addEventListener("webglcontextrestored", () => {
    webglLostAt = 0;
    showPlacementHint("3D context restored.", "ok");
    rebuildContainer();
    queueViewportRefresh(3);
    updateStats();
  });

  controls = new OrbitControls(camera, refs.canvas);
  controls.enableDamping = true;
  controls.target.set(0, 1, 0);

  raycaster = new THREE.Raycaster();

  const hemi = new THREE.HemisphereLight("#ffffff", "#bfd4ff", 0.95);
  const dir = new THREE.DirectionalLight("#ffffff", 0.75);
  dir.position.set(6, 12, 9);
  scene.add(hemi, dir);

  const grid = new THREE.GridHelper(16, 32, "#b7c4dc", "#dce4f1");
  grid.position.y = 0;
  scene.add(grid);

  scene.add(cargoGroup);

  rebuildContainer();
  resetMainView();
  onResize();
  queueViewportRefresh(2);
  window.addEventListener("resize", onResize);
  window.addEventListener("resize", () => {
    if (tourActive) positionTour();
  });
  window.addEventListener("scroll", () => {
    if (tourActive) positionTour();
  }, { passive: true });
  if (window.ResizeObserver && !viewportResizeObserver) {
    viewportResizeObserver = new ResizeObserver(() => {
      queueViewportRefresh(2);
      if (tourActive) positionTour();
    });
    viewportResizeObserver.observe(refs.viewportWrap);
  }
  animate();
};

const rebuildContainer = () => {
  if (containerGroup) {
    scene.remove(containerGroup);
    disposeObject3D(containerGroup);
  }

  const cfg = getContainerCfg();
  const length = cfg.lengthMm * MM_TO_M;
  const width = cfg.widthMm * MM_TO_M;
  const height = cfg.heightMm * MM_TO_M;
  const ribCount = cfg.model?.showRibs ? (cfg.model?.ribCount ?? (length > 8 ? 36 : 18)) : 0;
  const ribDepth = (cfg.model?.ribDepthMm ?? 24) * MM_TO_M;
  const cornerPost = (cfg.model?.cornerPostMm ?? 70) * MM_TO_M;
  const shellColor = cfg.model?.shellColor ?? "#dfe7f7";
  const edgeColor = cfg.model?.edgeColor ?? "#465b80";

  containerGroup = new THREE.Group();

  const boxGeom = new THREE.BoxGeometry(length, height, width);
  const shellMat = new THREE.MeshPhysicalMaterial({
    color: "#dbe9ff",
    transparent: true,
    opacity: 0.14,
    roughness: 0.06,
    metalness: 0.02,
    transmission: 0.9,
    ior: 1.46,
    thickness: 0.12,
    clearcoat: 1.0,
    clearcoatRoughness: 0.08,
    reflectivity: 0.55,
    envMapIntensity: 1.6,
    side: THREE.DoubleSide,
  });

  const shell = new THREE.Mesh(boxGeom, shellMat);
  shell.position.y = height / 2;

  const edges = new THREE.EdgesGeometry(boxGeom);
  const wireMat = new THREE.LineBasicMaterial({
    color: edgeColor,
    transparent: true,
    opacity: 0.68,
  });
  wireMesh = new THREE.LineSegments(edges, wireMat);
  wireMesh.position.y = height / 2;

  const floorGeom = new THREE.PlaneGeometry(length, width);
  const floorMat = new THREE.MeshStandardMaterial({
    color: "#f7fbff",
    roughness: 0.85,
    metalness: 0.04,
    side: THREE.DoubleSide,
  });
  floorMesh = new THREE.Mesh(floorGeom, floorMat);
  floorMesh.rotation.x = -Math.PI / 2;
  floorMesh.position.y = 0.001;

  if (ribCount > 0) {
    const ribMat = new THREE.MeshStandardMaterial({
      color: edgeColor,
      roughness: 0.35,
      metalness: 0.22,
    });
    const ribGeom = new THREE.BoxGeometry(length / ribCount * 0.12, height * 0.96, ribDepth);
    for (let i = 0; i <= ribCount; i += 1) {
      const t = i / ribCount;
      const x = -length / 2 + length * t;
      const leftRib = new THREE.Mesh(ribGeom, ribMat);
      leftRib.position.set(x, height / 2, -width / 2 + ribDepth / 2);
      const rightRib = new THREE.Mesh(ribGeom, ribMat);
      rightRib.position.set(x, height / 2, width / 2 - ribDepth / 2);
      containerGroup.add(leftRib, rightRib);
    }
  }

  const postGeom = new THREE.BoxGeometry(cornerPost, height, cornerPost);
  const postMat = new THREE.MeshStandardMaterial({
    color: edgeColor,
    roughness: 0.3,
    metalness: 0.28,
  });
  const postOffsets = [
    [-length / 2 + cornerPost / 2, height / 2, -width / 2 + cornerPost / 2],
    [-length / 2 + cornerPost / 2, height / 2, width / 2 - cornerPost / 2],
    [length / 2 - cornerPost / 2, height / 2, -width / 2 + cornerPost / 2],
    [length / 2 - cornerPost / 2, height / 2, width / 2 - cornerPost / 2],
  ];
  postOffsets.forEach(([x, y, z]) => {
    const post = new THREE.Mesh(postGeom, postMat);
    post.position.set(x, y, z);
    containerGroup.add(post);
  });

  const b = getContainerBounds();
  const usableLen = b.maxX - b.minX;
  const usableWid = b.maxZ - b.minZ;
  const usableHei = b.maxY;
  const usableGeom = new THREE.BoxGeometry(usableLen, usableHei, usableWid);
  const usableEdge = new THREE.EdgesGeometry(usableGeom);
  const usableMat = new THREE.LineBasicMaterial({ color: "#4da3ff", transparent: true, opacity: 0.35 });
  usableBoundsMesh = new THREE.LineSegments(usableEdge, usableMat);
  usableBoundsMesh.position.set((b.minX + b.maxX) / 2, usableHei / 2, (b.minZ + b.maxZ) / 2);

  containerGroup.add(shell, wireMesh, floorMesh, usableBoundsMesh);
  scene.add(containerGroup);

  controls.target.set(0, height * 0.4, 0);
  controls.update();
};

const onResize = () => {
  if (!renderer || !camera || !refs.viewportWrap) return;
  const rect = refs.viewportWrap.getBoundingClientRect();
  const width = Math.max(320, rect.width);
  const height = Math.max(320, rect.height);
  camera.aspect = width / height;
  camera.updateProjectionMatrix();
  renderer.setSize(width, height, false);
};

const animate = () => {
  requestAnimationFrame(animate);
  frameTick += 1;
  if (frameTick % 120 === 0) {
    ensureContainerVisible();
  }
  if (webglLostAt && performance.now() - webglLostAt > 420) {
    const lost = renderer?.getContext?.()?.isContextLost?.() || false;
    if (lost) {
      recoverMainRenderer("watchdog");
      webglLostAt = 0;
      return;
    }
  }
  try {
    controls.update();
    renderer.render(scene, camera);
    lastRenderAt = performance.now();
  } catch (err) {
    console.error("Render loop error:", err);
    recoverMainRenderer("render-error");
  }
};

const updateDropNdcFromPointer = (clientX, clientY, out) => {
  const rect = refs.canvas.getBoundingClientRect();
  out.x = ((clientX - rect.left) / rect.width) * 2 - 1;
  out.y = -((clientY - rect.top) / rect.height) * 2 + 1;
};

const getContainerBounds = () => {
  const cfg = getContainerCfg();
  const length = cfg.lengthMm * MM_TO_M;
  const width = cfg.widthMm * MM_TO_M;
  const height = cfg.heightMm * MM_TO_M;
  // Capacity-first bounds: only a tiny inset to avoid z-fighting with the rendered shell.
  // Structural posts are visualized externally and should not reduce planning capacity.
  const shellInset = 0.006;
  const insetX = shellInset;
  const insetZ = shellInset;
  const usableLength = Math.max(0.01, length - insetX * 2);
  const usableWidth = Math.max(0.01, width - insetZ * 2);
  return {
    minX: -length / 2 + insetX,
    maxX: length / 2 - insetX,
    minZ: -width / 2 + insetZ,
    maxZ: width / 2 - insetZ,
    maxY: height,
    floorArea: usableLength * usableWidth,
  };
};

const getOrientationOptions = (model) => {
  const [len, wid] = model.sizeM;
  const same = Math.abs(len - wid) < 1e-6;
  if (same) {
    return [{ len, wid, rotationY: 0 }];
  }
  return [
    { len, wid, rotationY: 0 },
    { len: wid, wid: len, rotationY: Math.PI / 2 },
  ];
};

const intersectsAny = (candidate) => {
  return state.placed.some((p) => {
    const xOverlap = candidate.x < p.x + p.len && candidate.x + candidate.len > p.x;
    const zOverlap = candidate.z < p.z + p.wid && candidate.z + candidate.wid > p.z;
    const yOverlap = candidate.y < p.y + p.hei && candidate.y + candidate.hei > p.y;
    return xOverlap && zOverlap && yOverlap;
  });
};

const inBounds = (candidate, bounds) => {
  const eps = 1e-6;
  return (
    candidate.x >= bounds.minX - eps &&
    candidate.y >= 0 &&
    candidate.z >= bounds.minZ - eps &&
    candidate.x + candidate.len <= bounds.maxX + eps &&
    candidate.z + candidate.wid <= bounds.maxZ + eps &&
    candidate.y + candidate.hei <= bounds.maxY + eps
  );
};

const canAddWeight = (model) => currentWeightKg() + model.palletWeightKg <= state.maxWeightKg;

const currentWeightKg = () => state.placed.reduce((sum, p) => sum + p.weightKg, 0);
const currentGrossKg = () => currentWeightKg() + getContainerCfg().tareKg;

const currentFootprintArea = () => state.placed.filter((p) => p.y === 0).reduce((sum, p) => sum + p.len * p.wid, 0);

const hasSupport = (candidate) => {
  if (candidate.y <= 0) return true;
  return state.placed.some((p) => {
    const topMatches = Math.abs(p.y + p.hei - candidate.y) < 1e-6;
    const xContains = candidate.x >= p.x && candidate.x + candidate.len <= p.x + p.len;
    const zContains = candidate.z >= p.z && candidate.z + candidate.wid <= p.z + p.wid;
    return topMatches && xContains && zContains;
  });
};

const findSupportingPlacement = (candidate) => {
  if (candidate.y <= 0) return null;
  return state.placed.find((p) => {
    const topMatches = Math.abs(p.y + p.hei - candidate.y) < 1e-6;
    const xContains = candidate.x >= p.x && candidate.x + candidate.len <= p.x + p.len;
    const zContains = candidate.z >= p.z && candidate.z + candidate.wid <= p.z + p.wid;
    return topMatches && xContains && zContains;
  }) || null;
};

const addPlacement = (model, x, y, z, options = {}) => {
  if (!canAddWeight(model)) return false;
  const [defaultLen, defaultWid, hei] = model.sizeM;
  const len = options.len ?? defaultLen;
  const wid = options.wid ?? defaultWid;
  const rotationY = options.rotationY ?? 0;
  const placementWeight = options.weightKg ?? model.palletWeightKg;
  const candidate = { x, y, z, len, wid, hei };
  const bounds = getContainerBounds();
  if (!inBounds(candidate, bounds) || intersectsAny(candidate) || !hasSupport(candidate)) return false;
  const support = findSupportingPlacement(candidate);
  if (support && placementWeight > support.weightKg) return false;

  const geom = new THREE.BoxGeometry(len, hei, wid);
  const mat = new THREE.MeshStandardMaterial({ color: model.color, roughness: 0.3, metalness: 0.05 });
  const mesh = new THREE.Mesh(geom, mat);
  mesh.position.set(x + len / 2, y + hei / 2, z + wid / 2);
  mesh.rotation.y = rotationY;

  const edge = new THREE.LineSegments(new THREE.EdgesGeometry(geom), new THREE.LineBasicMaterial({ color: "#1f2635" }));
  mesh.add(edge);

  const id = state.nextId++;
  mesh.userData.placementId = id;
  cargoGroup.add(mesh);

  state.placed.push({
    id,
    modelId: model.id,
    product: model.product,
    model: model.model,
    weightKg: placementWeight,
    x,
    y,
    z,
    len,
    wid,
    hei,
    rotationY,
    mesh,
  });

  invalidateContainerPlans();
  updateStats();
  return true;
};

const removePlacement = (placementId) => {
  const idx = state.placed.findIndex((p) => p.id === placementId);
  if (idx < 0) return;
  const [removed] = state.placed.splice(idx, 1);
  cargoGroup.remove(removed.mesh);
  disposeObject3D(removed.mesh);
  invalidateContainerPlans();
  updateStats();
};

const generateCandidateSpots = (preferred = null) => {
  const spots = [];
  if (preferred) spots.push(preferred);
  spots.push({ x: getContainerBounds().minX, y: 0, z: getContainerBounds().minZ });

  state.placed.forEach((p) => {
    spots.push({ x: p.x + p.len, y: p.y, z: p.z });
    spots.push({ x: p.x, y: p.y, z: p.z + p.wid });
    spots.push({ x: p.x + p.len, y: p.y, z: p.z + p.wid });
    spots.push({ x: p.x, y: p.y + p.hei, z: p.z });
  });

  const seen = new Set();
  return spots.filter((s) => {
    const key = `${s.x.toFixed(3)}_${s.y.toFixed(3)}_${s.z.toFixed(3)}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
};

const findFitSpot = (model, preferred = null) => {
  const bounds = getContainerBounds();
  const [, , hei] = model.sizeM;
  const orientations = getOrientationOptions(model);
  const containerLen = bounds.maxX - bounds.minX;
  const containerWid = bounds.maxZ - bounds.minZ;
  const candidates = generateCandidateSpots(preferred);

  // Add explicit stack anchors so mixed pallet sizes can be stacked on top supports.
  state.placed.forEach((p) => {
    const topY = p.y + p.hei;
    orientations.forEach((o) => {
      if (topY + hei > bounds.maxY) return;
      const topAnchors = [
        { x: p.x, y: topY, z: p.z, len: o.len, wid: o.wid, rotationY: o.rotationY },
        { x: p.x + p.len - o.len, y: topY, z: p.z, len: o.len, wid: o.wid, rotationY: o.rotationY },
        { x: p.x, y: topY, z: p.z + p.wid - o.wid, len: o.len, wid: o.wid, rotationY: o.rotationY },
        { x: p.x + p.len - o.len, y: topY, z: p.z + p.wid - o.wid, len: o.len, wid: o.wid, rotationY: o.rotationY },
        { x: p.x + (p.len - o.len) / 2, y: topY, z: p.z + (p.wid - o.wid) / 2, len: o.len, wid: o.wid, rotationY: o.rotationY },
      ];
      topAnchors.forEach((a) => {
        if (Number.isFinite(a.x) && Number.isFinite(a.z)) candidates.push(a);
      });
    });
  });

  candidates.sort((a, b) => (a.y - b.y) || (a.z - b.z) || (a.x - b.x));

  const feasible = [];
  for (const candidate of candidates) {
    const preferredOrientations = candidate.len && candidate.wid
      ? [{ len: candidate.len, wid: candidate.wid, rotationY: candidate.rotationY || 0 }]
      : orientations;
    for (const o of preferredOrientations) {
      const box = { x: candidate.x, y: candidate.y, z: candidate.z, len: o.len, wid: o.wid, hei };
      if (inBounds(box, bounds) && !intersectsAny(box) && hasSupport(box)) {
        const perLayerCapacity = Math.floor(containerLen / o.len) * Math.floor(containerWid / o.wid);
        const residualLen = containerLen - Math.floor(containerLen / o.len) * o.len;
        const residualWid = containerWid - Math.floor(containerWid / o.wid) * o.wid;
        feasible.push({
          x: candidate.x,
          y: candidate.y,
          z: candidate.z,
          len: o.len,
          wid: o.wid,
          rotationY: o.rotationY,
          perLayerCapacity,
          residualScore: residualLen + residualWid,
        });
      }
    }
  }
  if (!feasible.length) return null;
  feasible.sort((a, b) => {
    if (a.y !== b.y) return a.y - b.y;
    if (b.perLayerCapacity !== a.perLayerCapacity) return b.perLayerCapacity - a.perLayerCapacity;
    if (a.residualScore !== b.residualScore) return a.residualScore - b.residualScore;
    if (a.z !== b.z) return a.z - b.z;
    return a.x - b.x;
  });
  return feasible[0];
};

const placeModelSmart = (model, preferred = null) => {
  const spot = findFitSpot(model, preferred);
  if (!spot) return false;
  return addPlacement(model, spot.x, spot.y, spot.z, { len: spot.len, wid: spot.wid, rotationY: spot.rotationY });
};

const dropModelAtPointer = (model, ndc) => {
  raycaster.setFromCamera(ndc, camera);
  const hit = raycaster.intersectObject(floorMesh, false)[0];
  if (!hit) return false;
  const snap = 0.05;
  const orientations = getOrientationOptions(model);
  for (const o of orientations) {
    const x = Math.round((hit.point.x - o.len / 2) / snap) * snap;
    const z = Math.round((hit.point.z - o.wid / 2) / snap) * snap;
    if (addPlacement(model, x, 0, z, { len: o.len, wid: o.wid, rotationY: o.rotationY })) return true;
  }
  const [first] = orientations;
  const x = Math.round((hit.point.x - first.len / 2) / snap) * snap;
  const z = Math.round((hit.point.z - first.wid / 2) / snap) * snap;
  return placeModelSmart(model, { x, y: 0, z });
};

const fillByQuantity = () => {
  const tasks = Object.entries(state.qtyByModel)
    .map(([id, qty]) => ({ id, qty: Number(qty) || 0 }))
    .filter((item) => item.qty > 0);

  for (const task of tasks) {
    const model = PALLET_MODELS.find((m) => m.id === task.id);
    if (!model || !isModelEnabled(model)) continue;
    for (let i = 0; i < task.qty; i += 1) {
      if (!placeModelSmart(model)) return;
    }
  }
};

const autoFillGreedy = () => {
  const before = state.placed.length;
  let guard = 0;
  while (guard < 800) {
    guard += 1;
    const options = PALLET_MODELS
      .filter((m) => isModelEnabled(m) && canAddWeight(m))
      .map((m) => ({ model: m, spot: findFitSpot(m) }))
      .filter((entry) => entry.spot);

    if (!options.length) break;

    options.sort((a, b) => {
      const aScore = (a.model.sizeM[0] * a.model.sizeM[1]) / a.model.palletWeightKg;
      const bScore = (b.model.sizeM[0] * b.model.sizeM[1]) / b.model.palletWeightKg;
      return bScore - aScore;
    });

    const choice = options[0];
    if (!choice || !choice.spot) break;
    addPlacement(choice.model, choice.spot.x, choice.spot.y, choice.spot.z, {
      len: choice.spot.len,
      wid: choice.spot.wid,
      rotationY: choice.spot.rotationY,
    });
  }
  const added = state.placed.length - before;
  if (added > 0) {
    showPlacementHint(
      `Quick Fill loaded ${added} pallets (${state.placed.length} total). Demand quantities were ignored.`,
      "ok"
    );
  } else {
    showPlacementHint("Quick Fill could not place more pallets under current space/weight limits.", "warn");
  }
};

const preferredSpotFromNdc = (model, ndc) => {
  if (!ndc) return null;
  raycaster.setFromCamera(ndc, camera);
  const hit = raycaster.intersectObject(floorMesh, false)[0];
  if (!hit) return null;
  const [len, wid] = model.sizeM;
  const snap = 0.05;
  const x = Math.round((hit.point.x - len / 2) / snap) * snap;
  const z = Math.round((hit.point.z - wid / 2) / snap) * snap;
  return { x, y: 0, z };
};

const scorePalletCandidate = (model, demand, constrainedByWeight) => {
  const units = model.units;
  const fulfill = Math.min(demand, units);
  const overshoot = Math.max(0, units - demand);
  const footprint = Math.max(1e-6, model.sizeM[0] * model.sizeM[1]);
  const unitsPerArea = units / footprint;
  const unitsPerWeight = units / Math.max(1e-6, model.palletWeightKg);
  const unitsPerVolume = units / Math.max(1e-6, model.sizeM[0] * model.sizeM[1] * model.sizeM[2]);

  const potentialUnits = getModelCapacityPotential(model).units;
  const demandFitRatio = demand > 0 ? fulfill / demand : 0;
  const demandRemainderPenalty = demand > units ? 0 : overshoot;

  if (state.planStrategy === "space_first") {
    return (potentialUnits * 28)
      + (unitsPerArea * 1250)
      + (unitsPerVolume * 520)
      + (demandFitRatio * 240)
      - (demandRemainderPenalty * 120);
  }
  if (state.planStrategy === "min_pallets") {
    return (potentialUnits * 18)
      + (unitsPerWeight * 520)
      + (unitsPerVolume * 420)
      + (demandFitRatio * 200)
      - (demandRemainderPenalty * 80);
  }

  const efficiency = constrainedByWeight ? unitsPerWeight : unitsPerArea;
  return (potentialUnits * 32)
    + (efficiency * 560)
    + (unitsPerVolume * 680)
    + (demandFitRatio * 220)
    - (demandRemainderPenalty * 110);
};

const chooseBestPalletCandidate = (product, demand = 1, preferredNdc = null) => {
  const bounds = getContainerBounds();
  const floorArea = Math.max(bounds.floorArea, 1e-6);
  const weightRate = state.maxWeightKg > 0 ? currentWeightKg() / state.maxWeightKg : 0;
  const footprintRate = currentFootprintArea() / floorArea;
  const constrainedByWeight = weightRate >= footprintRate;

  const candidates = PALLET_MODELS
    .filter((model) => model.product === product && isModelEnabled(model) && canAddWeight(model))
    .map((model) => {
      const preferred = preferredSpotFromNdc(model, preferredNdc);
      const spot = findFitSpot(model, preferred);
      if (!spot) return null;
      return {
        model,
        spot,
        score: scorePalletCandidate(model, demand, constrainedByWeight),
        units: model.units,
      };
    })
    .filter(Boolean);

  if (!candidates.length) return null;
  candidates.sort((a, b) => {
    if (b.score !== a.score) return b.score - a.score;
    return b.units - a.units;
  });
  return candidates[0];
};

const getModelCapacityPotential = (model) => {
  const bounds = getContainerBounds();
  const layerCount = Math.max(1, Math.floor(bounds.maxY / model.sizeM[2]));
  let bestBySpace = 0;
  getOrientationOptions(model).forEach((o) => {
    const lenCount = Math.floor((bounds.maxX - bounds.minX) / o.len);
    const widCount = Math.floor((bounds.maxZ - bounds.minZ) / o.wid);
    bestBySpace = Math.max(bestBySpace, lenCount * widCount * layerCount);
  });
  const bySpacePallets = bestBySpace;
  const byWeightPallets = Math.max(0, Math.floor(state.maxWeightKg / Math.max(1e-6, model.palletWeightKg)));
  return {
    pallets: Math.min(bySpacePallets, byWeightPallets),
    units: Math.min(bySpacePallets, byWeightPallets) * model.units,
    layers: layerCount,
  };
};

const getStackCapacityPotential = (spot, stackUnits, stackWeightKg) => {
  const bounds = getContainerBounds();
  const lenCount = Math.max(1, Math.floor((bounds.maxX - bounds.minX) / Math.max(1e-6, spot.len)));
  const widCount = Math.max(1, Math.floor((bounds.maxZ - bounds.minZ) / Math.max(1e-6, spot.wid)));
  const bySpaceColumns = lenCount * widCount;
  const byWeightColumns = Math.max(0, Math.floor(state.maxWeightKg / Math.max(1e-6, stackWeightKg)));
  const columns = Math.min(bySpaceColumns, byWeightColumns);
  return {
    columns,
    units: columns * stackUnits,
  };
};

const buildBestStackPlanForBase = (product, baseModel, baseSpot) => {
  const bounds = getContainerBounds();
  const basePlacement = {
    model: baseModel,
    placement: {
      x: baseSpot.x,
      y: baseSpot.y,
      z: baseSpot.z,
      len: baseSpot.len,
      wid: baseSpot.wid,
      rotationY: baseSpot.rotationY || 0,
    },
  };
  const maxWeightAllowance = Math.max(0, state.maxWeightKg - currentWeightKg());
  const baseWeight = baseModel.palletWeightKg;
  if (baseWeight > maxWeightAllowance + 1e-6) {
    return { placements: [basePlacement], units: baseModel.units, weightKg: baseWeight };
  }

  const topOptions = PALLET_MODELS
    .filter((m) => m.product === product && isModelEnabled(m))
    .flatMap((m) =>
      getOrientationOptions(m)
        .filter((o) => o.len <= baseSpot.len + 1e-6 && o.wid <= baseSpot.wid + 1e-6)
        .map((o) => ({ model: m, len: o.len, wid: o.wid, rotationY: o.rotationY }))
    );

  const maxLayers = 6;
  const best = {
    placements: [basePlacement],
    units: baseModel.units,
    weightKg: baseWeight,
  };

  const dfs = (chain, nextY, lastLen, lastWid, lastWeightKg, units, weightKg) => {
    if (units > best.units || (units === best.units && weightKg < best.weightKg)) {
      best.placements = chain.map((entry) => ({
        model: entry.model,
        placement: { ...entry.placement },
      }));
      best.units = units;
      best.weightKg = weightKg;
    }
    if (chain.length >= maxLayers) return;

    topOptions.forEach((option) => {
      if (option.model.palletWeightKg > lastWeightKg + 1e-6) return;
      if (option.len > lastLen + 1e-6 || option.wid > lastWid + 1e-6) return;
      const hei = option.model.sizeM[2];
      if (nextY + hei > bounds.maxY + 1e-6) return;
      const nextWeight = weightKg + option.model.palletWeightKg;
      if (nextWeight > maxWeightAllowance + 1e-6) return;

      const candidateBox = {
        x: baseSpot.x,
        y: nextY,
        z: baseSpot.z,
        len: option.len,
        wid: option.wid,
        hei,
      };
      if (!inBounds(candidateBox, bounds)) return;

      const nextEntry = {
        model: option.model,
        placement: {
          x: baseSpot.x,
          y: nextY,
          z: baseSpot.z,
          len: option.len,
          wid: option.wid,
          rotationY: option.rotationY,
        },
      };

      chain.push(nextEntry);
      dfs(
        chain,
        nextY + hei,
        option.len,
        option.wid,
        option.model.palletWeightKg,
        units + option.model.units,
        nextWeight
      );
      chain.pop();
    });
  };

  dfs(
    [basePlacement],
    baseSpot.y + baseModel.sizeM[2],
    baseSpot.len,
    baseSpot.wid,
    baseModel.palletWeightKg,
    baseModel.units,
    baseWeight
  );
  return best;
};

const listPalletCandidatesForProduct = (product, demand = 1, preferredNdc = null) => {
  const bounds = getContainerBounds();
  const floorArea = Math.max(bounds.floorArea, 1e-6);
  const weightRate = state.maxWeightKg > 0 ? currentWeightKg() / state.maxWeightKg : 0;
  const footprintRate = currentFootprintArea() / floorArea;
  const constrainedByWeight = weightRate >= footprintRate;

  return PALLET_MODELS
    .filter((model) => model.product === product && isModelEnabled(model) && canAddWeight(model))
    .map((model) => {
      const preferred = preferredSpotFromNdc(model, preferredNdc);
      const spot = findFitSpot(model, preferred);
      if (!spot) return null;
      const stackPlan = spot.y === 0
        ? buildBestStackPlanForBase(product, model, spot)
        : {
            placements: [{
              model,
              placement: {
                x: spot.x,
                y: spot.y,
                z: spot.z,
                len: spot.len,
                wid: spot.wid,
                rotationY: spot.rotationY || 0,
              },
            }],
            units: model.units,
            weightKg: model.palletWeightKg,
          };
      const stackUnits = stackPlan.units;
      const stackWeightKg = stackPlan.weightKg;
      const fulfill = Math.min(demand, stackUnits);
      const overshoot = Math.max(0, stackUnits - demand);
      const potential = getStackCapacityPotential(spot, stackUnits, stackWeightKg);
      const basePotential = getModelCapacityPotential(model);

      return {
        model,
        spot,
        score: scorePalletCandidate(model, demand, constrainedByWeight),
        fulfill,
        overshoot,
        stackPlan: stackPlan.placements,
        stackUnits,
        stackWeightKg,
        stackPallets: stackPlan.placements.length,
        potentialUnits: potential.units,
        potentialPallets: potential.columns,
        potentialLayers: basePotential.layers,
      };
    })
    .filter(Boolean);
};

const placeBestPalletForProduct = (product, preferredNdc = null) => {
  if (!state.enabledProducts.has(product)) return false;
  const demand = sanitizeDemand(state.productDemand[product]) || 1;
  const choice = chooseBestPalletCandidate(product, demand, preferredNdc);
  if (!choice) return false;
  const ok = addPlacement(choice.model, choice.spot.x, choice.spot.y, choice.spot.z, {
    len: choice.spot.len,
    wid: choice.spot.wid,
    rotationY: choice.spot.rotationY,
  });
  if (!ok) return false;
  state.activeProduct = product;
  renderProductMix();
  renderPalletCards();
  showPlacementHint(
    `Loaded by product: ${choice.model.product} [Pallet ${choice.model.model}] · ${state.placed.length} pallets`,
    "ok"
  );
  return true;
};

const applyBundleAndReplan = (bundleId, qty = 1) => {
  const bundle = PRODUCT_BUNDLES.find((b) => b.id === bundleId);
  if (!bundle) return false;
  const packQty = Math.max(1, sanitizeDemand(qty) || 1);
  bundle.items.forEach((item) => {
    state.enabledProducts.add(item.product);
    state.productDemand[item.product] = sanitizeDemand(state.productDemand[item.product]) + (item.units * packQty);
  });
  if (!state.activeProduct || !state.enabledProducts.has(state.activeProduct)) {
    state.activeProduct = bundle.items[0]?.product || getFirstEnabledProduct();
  }
  renderProductMix();
  renderBundleLibrary();
  return autoPlanByProductDemand();
};

const autoPlanByProductDemand = () => {
  clearAlgoTrace();
  ensureProductDemand();
  const enabledProducts = getEnabledProductList();
  const remaining = new Map(
    enabledProducts
      .map((product) => [product, sanitizeDemand(state.productDemand[product])])
      .filter(([, demand]) => demand > 0)
  );

  if (!remaining.size) {
    logAlgoTrace("No demand input found. Planning stopped.", "warn");
    showPlacementHint("Set Demand pieces first, then run Plan by Demand.", "warn");
    return false;
  }
  logAlgoTrace(`Plan started for ${enabledProducts.length} product(s) in ${getContainerCfg().label}.`, "info");

  state.isBatchPlanning = true;
  state.containerPlans = [];
  state.activeContainerPlan = 0;
  state.qtyByModel = {};
  clearActivePlacements(false);

  const maxContainers = 24;
  let containerGuard = 0;

  while (containerGuard < maxContainers) {
    containerGuard += 1;
    clearActivePlacements(false);
    const containerModelUsage = {};
    logAlgoTrace(`Container #${containerGuard}: planning phase started.`, "info");

    let guard = 0;
    const maxRounds = 1500;

    while (guard < maxRounds) {
      guard += 1;
      const needs = Array.from(remaining.entries()).filter(([, qty]) => qty > 0);
      const inDemandPhase = needs.length > 0;
      if (guard === 1) {
        logAlgoTrace(inDemandPhase ? `Container #${containerGuard}: demand-satisfaction phase.` : `Container #${containerGuard}: value-fill phase.`, "info");
      }

      const candidates = [];
      if (inDemandPhase) {
        // Phase 1: satisfy minimum demand first, evaluating A/B/C alternatives together.
        needs.forEach(([product, demand]) => {
          const productCandidates = listPalletCandidatesForProduct(product, demand);
          productCandidates.forEach((entry) => {
            candidates.push({ ...entry, inDemandPhase: true });
          });
        });
      } else {
        // Phase 2: demand satisfied; evaluate mixed pallet combinations to maximize value.
        const fillProducts = getEnabledProductList();
        fillProducts.forEach((product) => {
          const productCandidates = listPalletCandidatesForProduct(product, 999999);
          productCandidates.forEach((entry) => {
            const footprint = Math.max(1e-6, entry.spot.len * entry.spot.wid);
            const stackHeight = entry.stackPlan.reduce((sum, layer) => sum + layer.model.sizeM[2], 0);
            const unitsPerArea = entry.stackUnits / footprint;
            const unitsPerWeight = entry.stackUnits / Math.max(1e-6, entry.stackWeightKg);
            const unitsPerVolume = entry.stackUnits / Math.max(1e-6, footprint * Math.max(1e-6, stackHeight));
            const bounds = getContainerBounds();
            const lenCount = Math.max(1, Math.floor((bounds.maxX - bounds.minX) / entry.spot.len));
            const widCount = Math.max(1, Math.floor((bounds.maxZ - bounds.minZ) / entry.spot.wid));
            const bySpaceCols = lenCount * widCount;
            const byWeightCols = Math.max(0, Math.floor(state.maxWeightKg / Math.max(1e-6, entry.stackWeightKg)));
            const potentialUnits = Math.min(bySpaceCols, byWeightCols) * entry.stackUnits;
            candidates.push({
              ...entry,
              inDemandPhase: false,
              fillScore: (potentialUnits * 34)
                + (unitsPerVolume * 820)
                + (unitsPerArea * 260)
                + (unitsPerWeight * 220)
                + (entry.stackUnits * 700),
            });
          });
        });
      }

      if (!candidates.length) break;

      const usageByProduct = {};
      Object.entries(containerModelUsage).forEach(([modelId, count]) => {
        const model = PALLET_MODELS.find((m) => m.id === modelId);
        if (!model) return;
        usageByProduct[model.product] = (usageByProduct[model.product] || 0) + count;
      });

      candidates.forEach((entry) => {
        const usedThisModel = containerModelUsage[entry.model.id] || 0;
        const usedThisProduct = usageByProduct[entry.model.product] || 0;
        const usedShare = usedThisProduct > 0 ? usedThisModel / usedThisProduct : 0;
        const hasAltForSameProduct = candidates.some(
          (c) => c.model.product === entry.model.product && c.model.id !== entry.model.id
        );
        const dynamicMixBonus = (usedThisModel === 0 ? 180 : 0) - (usedThisModel * 32) - (hasAltForSameProduct ? usedShare * 260 : 0);
        const comboBonus = Math.max(0, entry.stackUnits - entry.model.units) * 420;

        if (entry.inDemandPhase) {
          entry.priority =
            (entry.potentialUnits * 2200) +
            (entry.stackUnits * 900) +
            (entry.fulfill * 260) -
            (entry.overshoot * 110) +
            (entry.score * 1.1) +
            dynamicMixBonus +
            comboBonus;
        } else {
          entry.priority =
            ((entry.fillScore || 0) * 95) +
            (entry.potentialUnits * 1500) +
            (entry.stackUnits * 1100) +
            (entry.score * 0.4) +
            (entry.spot.y > 0 ? 24 : 0) +
            dynamicMixBonus +
            comboBonus;
        }
      });

      candidates.sort((a, b) => {
        if (b.priority !== a.priority) return b.priority - a.priority;
        if (b.potentialUnits !== a.potentialUnits) return b.potentialUnits - a.potentialUnits;
        return b.score - a.score;
      });

      const choice = candidates[0];
      if (!choice) {
        break;
      }

      const beforeSnap = snapshotCurrentPlacements();
      const stackPlan = choice.stackPlan?.length
        ? choice.stackPlan
        : [{
            model: choice.model,
            placement: {
              x: choice.spot.x,
              y: choice.spot.y,
              z: choice.spot.z,
              len: choice.spot.len,
              wid: choice.spot.wid,
              rotationY: choice.spot.rotationY || 0,
            },
          }];

      let stackApplied = true;
      for (let i = 0; i < stackPlan.length; i += 1) {
        const layer = stackPlan[i];
        const p = layer.placement;
        const ok = addPlacement(layer.model, p.x, p.y, p.z, {
          len: p.len,
          wid: p.wid,
          rotationY: p.rotationY || 0,
        });
        if (!ok) {
          stackApplied = false;
          break;
        }
      }

      if (!stackApplied) {
        clearActivePlacements(false);
        beforeSnap.forEach((snap) => restorePlacementSnapshot(snap));
        updateStats();
        break;
      }

      stackPlan.forEach((layer, idx) => {
        const m = layer.model;
        containerModelUsage[m.id] = (containerModelUsage[m.id] || 0) + 1;
        state.qtyByModel[m.id] = (state.qtyByModel[m.id] || 0) + 1;
        remaining.set(m.product, Math.max(0, (remaining.get(m.product) || 0) - m.units));
        const y = layer.placement?.y ?? 0;
        if (idx === 0) {
          logAlgoTrace(
            `C${containerGuard} · ${m.product} -> Pallet ${m.model} selected (units ${m.units}, y=${y.toFixed(2)}m).`,
            "success"
          );
        } else {
          logAlgoTrace(
            `C${containerGuard} · stacked Pallet ${m.model} (units ${m.units}, y=${y.toFixed(2)}m, rot=${Math.round(((layer.placement?.rotationY || 0) * 180) / Math.PI)}deg).`,
            "success"
          );
        }
      });
      if (stackPlan.length > 1) {
        const stackUnits = stackPlan.reduce((sum, layer) => sum + layer.model.units, 0);
        logAlgoTrace(`C${containerGuard} · stack optimized: ${stackPlan.length} pallets / ${stackUnits} units in one column.`, "info");
      }
    }

    if (!state.placed.length) break;

    const baseUnits = state.placed.reduce((sum, p) => {
      const model = PALLET_MODELS.find((m) => m.id === p.modelId);
      return sum + (model?.units || 0);
    }, 0);
    state.containerPlans.push({
      placements: snapshotCurrentPlacements(),
      pallets: state.placed.length,
      baseUnits,
    });
    logAlgoTrace(`Container #${containerGuard} complete: ${baseUnits} units / ${state.placed.length} pallets.`, "success");

    const stillNeed = Array.from(remaining.values()).some((qty) => qty > 0);
    if (!stillNeed) break;
  }

  state.isBatchPlanning = false;

  const plannedProducts = new Set(
    Object.keys(state.qtyByModel)
      .map((modelId) => PALLET_MODELS.find((m) => m.id === modelId)?.product)
      .filter(Boolean)
  );
  if (plannedProducts.size) {
    state.activeProduct = Array.from(plannedProducts)[0];
  }
  renderProductMix();
  renderPalletCards();

  if (!state.containerPlans.length) {
    clearActivePlacements(true);
    renderContainerTabs();
    logAlgoTrace("No feasible placement found under current constraints.", "warn");
    showPlacementHint("Plan by Demand could not place cargo under current space/weight limits.", "warn");
    return false;
  }

  loadContainerPlan(0);

  const unmetUnits = Array.from(remaining.values()).reduce((sum, qty) => sum + qty, 0);
  const loadedUnits = state.containerPlans.reduce((sum, plan) => sum + plan.baseUnits, 0);
  const totalPallets = state.containerPlans.reduce((sum, plan) => sum + plan.pallets, 0);
  const containerCount = state.containerPlans.length;

  if (unmetUnits > 0) {
    logAlgoTrace(`Planning partial. Unmet units: ${unmetUnits}.`, "warn");
    showPlacementHint(
      `Planned ${containerCount} containers: ${loadedUnits} units, ${totalPallets} pallets. ${unmetUnits} units still unmet.`,
      "warn"
    );
    return true;
  }

  logAlgoTrace(`Planning complete. ${loadedUnits} units in ${containerCount} container(s).`, "success");
  showPlacementHint(
    `Plan complete: minimum demand satisfied, then container filled to max value. ${loadedUnits} units in ${containerCount} container(s), ${totalPallets} pallets total.`,
    "ok"
  );
  return true;
};

const clearAllCargo = () => {
  state.placed.forEach((p) => {
    cargoGroup.remove(p.mesh);
    disposeObject3D(p.mesh);
  });
  state.placed = [];

  while (cargoGroup.children.length) {
    const child = cargoGroup.children[0];
    cargoGroup.remove(child);
    disposeObject3D(child);
  }

  state.selectedModelId = null;
  state.qtyByModel = {};
  state.containerPlans = [];
  state.activeContainerPlan = 0;
  renderContainerTabs();
  if (refs.activePlacement) refs.activePlacement.textContent = "Current Placement: None";
  renderPalletCards();

  updateStats();
};

const onViewportPointerDown = (event) => {
  updateDropNdcFromPointer(event.clientX, event.clientY, pointerNdc);
  raycaster.setFromCamera(pointerNdc, camera);

  const meshes = state.placed.map((p) => p.mesh);
  const hits = raycaster.intersectObjects(meshes, false);
  if (hits.length) {
    const placementId = hits[0].object.userData.placementId;
    if (placementId) {
      removePlacement(placementId);
      return;
    }
  }

  if (state.selectedModelId) {
    const model = PALLET_MODELS.find((m) => m.id === state.selectedModelId);
    if (model && isModelEnabled(model)) {
      let ok = dropModelAtPointer(model, pointerNdc);
      if (!ok) ok = placeModelSmart(model);
      if (ok) {
        showPlacementHint(
          `Loaded: ${model.product} ${model.model} · ${state.placed.length} pallets · ${Math.round(currentWeightKg())} kg`,
          "ok"
        );
      } else {
        showPlacementHint(`Cannot place: no space or payload left for ${model.product} ${model.model}`, "warn");
      }
    }
  }
};

const updateStats = () => {
  const total = state.placed.length;
  const weight = currentWeightKg();
  const gross = currentGrossKg();
  const util = state.maxWeightKg > 0 ? (weight / state.maxWeightKg) * 100 : 0;
  const bounds = getContainerBounds();
  const footprint = bounds.floorArea > 0 ? (currentFootprintArea() / bounds.floorArea) * 100 : 0;
  const baseUnits = state.placed.reduce((sum, p) => {
    const model = PALLET_MODELS.find((m) => m.id === p.modelId);
    return sum + (model ? model.units : 0);
  }, 0);
  const legendCounts = new Map();
  const productUnits = new Map();
  state.placed.forEach((p) => {
    const model = PALLET_MODELS.find((m) => m.id === p.modelId);
    if (!model) return;
    productUnits.set(model.product, (productUnits.get(model.product) || 0) + model.units);
    const key = `${model.id}`;
    legendCounts.set(key, {
      modelId: model.id,
      product: model.product,
      type: model.model,
      color: model.color,
      pallets: (legendCounts.get(key)?.pallets || 0) + 1,
    });
  });

  refs.statPallets.textContent = String(total);
  refs.statWeight.textContent = `${Math.round(weight)} kg (gross ${Math.round(gross)} kg)`;
  refs.statWeightUtil.textContent = `${util.toFixed(1)}%`;
  refs.statFootprint.textContent = `${Math.min(100, footprint).toFixed(1)}%`;
  if (refs.loadBaseUnits) refs.loadBaseUnits.textContent = String(baseUnits);
  if (refs.loadPalletCount) refs.loadPalletCount.textContent = String(total);
  if (refs.loadRate) refs.loadRate.textContent = `${Math.min(100, footprint).toFixed(1)}%`;
  if (refs.weightRate) refs.weightRate.textContent = `${Math.min(100, util).toFixed(1)}%`;
  if (refs.loadRateBar) refs.loadRateBar.style.width = `${Math.min(100, Math.max(0, footprint)).toFixed(1)}%`;
  if (refs.weightRateBar) refs.weightRateBar.style.width = `${Math.min(100, Math.max(0, util)).toFixed(1)}%`;
  if (refs.colorLegend) {
    if (legendCounts.size === 0) {
      refs.colorLegend.innerHTML = `<div class="legend-empty">Color legend appears here after loading pallets.</div>`;
    } else {
      refs.colorLegend.innerHTML = "";
      Array.from(legendCounts.values())
        .sort((a, b) => b.pallets - a.pallets)
        .forEach((entry) => {
          const row = document.createElement("div");
          row.className = "legend-row";
          row.innerHTML = `
            <span class="legend-swatch" style="background:${entry.color}"></span>
            <span class="legend-label">${entry.product} [Pallet Type ${entry.type}]</span>
            <span class="legend-count">${entry.pallets}</span>
          `;
          refs.colorLegend.appendChild(row);
        });
    }
  }
  if (refs.loadProductsList) {
    if (productUnits.size === 0) {
      refs.loadProductsList.innerHTML = `<div class="load-product-row"><span>No products loaded</span><strong>0</strong></div>`;
    } else {
      refs.loadProductsList.innerHTML = "";
      Array.from(productUnits.entries())
        .sort((a, b) => b[1] - a[1])
        .forEach(([name, units]) => {
          const row = document.createElement("div");
          row.className = "load-product-row";
          row.innerHTML = `<span>${name}</span><strong>${units}</strong>`;
          refs.loadProductsList.appendChild(row);
        });
    }
  }

  const grouped = new Map();
  state.placed.forEach((p) => {
    const key = `${p.product}|${p.model}`;
    const prev = grouped.get(key) || { count: 0, weight: 0, modelId: p.modelId };
    prev.count += 1;
    prev.weight += p.weightKg;
    grouped.set(key, prev);
  });

  refs.summaryTable.innerHTML = "";
  if (!grouped.size) {
    refs.summaryTable.innerHTML = '<div class="summary-row"><span></span><span>No pallets placed</span><span class="tag">0</span><span class="tag">0 kg</span></div>';
    renderComplianceChecks();
    return;
  }

  Array.from(grouped.entries()).forEach(([key, value]) => {
    const [product, model] = key.split("|");
    const modelInfo = PALLET_MODELS.find((m) => m.id === value.modelId);
    const row = document.createElement("div");
    row.className = "summary-row";
    row.innerHTML = `
      <span class="swatch" style="background:${modelInfo ? modelInfo.color : "#999"}"></span>
      <span>${product} ${model}</span>
      <span class="tag">${value.count} pallets</span>
      <span class="tag">${Math.round(value.weight)} kg</span>
    `;
    refs.summaryTable.appendChild(row);
  });

  renderComplianceChecks();
};

const overlapsOnAxis = (aMin, aMax, bMin, bMax) => aMin < bMax && aMax > bMin;

const touchesWall = (p, b) => (
  Math.abs(p.x - b.minX) <= TOUCH_TOL ||
  Math.abs((p.x + p.len) - b.maxX) <= TOUCH_TOL ||
  Math.abs(p.z - b.minZ) <= TOUCH_TOL ||
  Math.abs((p.z + p.wid) - b.maxZ) <= TOUCH_TOL
);

const hasNeighborContact = (p) => {
  return state.placed.some((q) => {
    if (p.id === q.id || p.y !== q.y) return false;
    const zOverlap = overlapsOnAxis(p.z, p.z + p.wid, q.z, q.z + q.wid);
    const xOverlap = overlapsOnAxis(p.x, p.x + p.len, q.x, q.x + q.len);
    const touchX = Math.abs((p.x + p.len) - q.x) <= TOUCH_TOL || Math.abs((q.x + q.len) - p.x) <= TOUCH_TOL;
    const touchZ = Math.abs((p.z + p.wid) - q.z) <= TOUCH_TOL || Math.abs((q.z + q.wid) - p.z) <= TOUCH_TOL;
    return (zOverlap && touchX) || (xOverlap && touchZ);
  });
};

const buildComplianceChecks = () => {
  const cfg = getContainerCfg();
  const payload = currentWeightKg();
  const gross = currentGrossKg();
  const b = getContainerBounds();
  const checks = [];

  checks.push({
    rule: "CSC Convention",
    status: state.cscPlateOk ? "pass" : "fail",
    text: state.cscPlateOk ? "CSC plate verification is confirmed." : "CSC plate verification is missing.",
  });
  checks.push({
    rule: "CSC Convention",
    status: state.cscExamOk ? "pass" : "fail",
    text: state.cscExamOk ? "Periodic examination status is valid." : "Periodic examination status is not valid.",
  });

  const payloadLimit = Math.min(state.maxWeightKg, cfg.maxPayloadKg);
  const payloadOk = payload <= payloadLimit;
  const grossOk = gross <= cfg.maxGrossKg;
  checks.push({
    rule: "ISO 668 / CSC",
    status: payloadOk && grossOk ? "pass" : "fail",
    text: `Payload ${Math.round(payload)} / ${Math.round(payloadLimit)} kg, gross ${Math.round(gross)} / ${Math.round(cfg.maxGrossKg)} kg.`,
  });

  const allInside = state.placed.every((p) => inBounds(p, b));
  checks.push({
    rule: "ISO 1496 / ISO 668",
    status: allInside ? "pass" : "fail",
    text: allInside ? "All pallets are within internal container dimensions." : "Some pallets exceed container dimensional limits.",
  });

  if (payload > 0) {
    const length = cfg.lengthMm * MM_TO_M;
    const cogX = state.placed.reduce((sum, p) => sum + (p.x + p.len / 2) * p.weightKg, 0) / payload;
    const cogLimit = Math.min(0.1 * length, 0.8);
    checks.push({
      rule: "CTU Code",
      status: Math.abs(cogX) <= cogLimit ? "pass" : "fail",
      text: `Longitudinal center-of-mass offset is ${cogX.toFixed(2)} m (limit ±${cogLimit.toFixed(2)} m).`,
    });

    const leftMass = state.placed.filter((p) => (p.z + p.wid / 2) < 0).reduce((s, p) => s + p.weightKg, 0);
    const rightMass = payload - leftMass;
    const heavySideRatio = Math.max(leftMass, rightMass) / payload;
    checks.push({
      rule: "CTU Code",
      status: heavySideRatio <= 0.6 ? "pass" : "warn",
      text: `Transverse balance: heavy side ${(heavySideRatio * 100).toFixed(1)}% of payload (target <= 60%).`,
    });
  } else {
    checks.push({ rule: "CTU Code", status: "warn", text: "No cargo loaded yet; distribution checks are pending." });
    checks.push({ rule: "CTU Code", status: "warn", text: "No cargo loaded yet; transverse balance is pending." });
  }

  const unstable = state.placed.filter((p) => p.y > 0).filter((p) => {
    const support = findSupportingPlacement(p);
    return support && p.weightKg > support.weightKg;
  }).length;
  checks.push({
    rule: "CTU Code",
    status: unstable === 0 ? "pass" : "fail",
    text: unstable === 0 ? "Stacking is stable (no heavier pallet on lighter support)." : `${unstable} stacked pallets are heavier than their support base.`,
  });

  const floorPallets = state.placed.filter((p) => p.y === 0);
  const unsecured = floorPallets.filter((p) => !touchesWall(p, b) && !hasNeighborContact(p)).length;
  const footprintUsage = b.floorArea > 0 ? currentFootprintArea() / b.floorArea : 0;
  checks.push({
    rule: "CTU Code",
    status: unsecured > 0 || (state.placed.length > 0 && footprintUsage < 0.55) ? "warn" : "pass",
    text: unsecured > 0
      ? `${unsecured} floor pallets appear free-standing; blocking/lashing is required.`
      : `Floor coverage ${(footprintUsage * 100).toFixed(1)}% (low coverage may need additional securing).`,
  });
  return checks;
};

const renderComplianceChecks = () => {
  const checks = buildComplianceChecks();
  state.lastComplianceChecks = checks;

  refs.complianceList.innerHTML = "";
  checks.forEach((c) => {
    const row = document.createElement("div");
    row.className = "compliance-item";
    row.innerHTML = `
      <span class="compliance-pill ${c.status}">${c.status.toUpperCase()}</span>
      <div class="compliance-text"><strong>${c.rule}:</strong> ${c.text}</div>
    `;
    refs.complianceList.appendChild(row);
  });
};

const formatNow = () => {
  const d = new Date();
  const pad = (v) => String(v).padStart(2, "0");
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
};

const buildPdfReportData = () => {
  const cfg = getContainerCfg();
  const payload = currentWeightKg();
  const gross = currentGrossKg();
  const floorArea = Math.max(1e-6, getContainerBounds().floorArea);
  const loadRate = (currentFootprintArea() / floorArea) * 100;
  const weightRate = state.maxWeightKg > 0 ? (payload / state.maxWeightKg) * 100 : 0;

  const grouped = new Map();
  const productsLoaded = new Map();
  state.placed.forEach((p) => {
    const m = PALLET_MODELS.find((x) => x.id === p.modelId);
    if (!m) return;
    const key = m.id;
    const prev = grouped.get(key) || {
      modelId: m.id,
      product: m.product,
      model: m.model,
      pallets: 0,
      units: 0,
      weightKg: 0,
      size: m.size,
      palletWeightKg: m.palletWeightKg,
      unitsPerPallet: m.units,
    };
    prev.pallets += 1;
    prev.units += m.units;
    prev.weightKg += p.weightKg;
    grouped.set(key, prev);
    productsLoaded.set(m.product, (productsLoaded.get(m.product) || 0) + m.units);
  });

  const selectedProducts = ALL_PRODUCTS
    .filter((p) => state.enabledProducts.has(p))
    .map((product) => {
      const demand = sanitizeDemand(state.productDemand[product]);
      const loadedUnits = productsLoaded.get(product) || 0;
      return {
        product,
        demand,
        loadedUnits,
        imageUrl: productImageUrl(product),
      };
    })
    .filter((row) => row.demand > 0 || row.loadedUnits > 0);

  const palletRows = Array.from(grouped.values()).sort((a, b) => b.units - a.units);

  return {
    generatedAt: formatNow(),
    container: cfg,
    containerType: state.containerKey,
    containerIndex: state.containerPlans.length > 0 ? state.activeContainerPlan + 1 : 1,
    totalContainers: Math.max(1, state.containerPlans.length || 1),
    pallets: palletRows.reduce((sum, row) => sum + row.pallets, 0),
    units: palletRows.reduce((sum, row) => sum + row.units, 0),
    payloadKg: Math.round(payload),
    grossKg: Math.round(gross),
    loadRate: Math.min(100, Math.max(0, loadRate)),
    weightRate: Math.min(100, Math.max(0, weightRate)),
    selectedProducts,
    palletRows,
  };
};

const loadImageDataUrl = (url) =>
  new Promise((resolve) => {
    if (!url) {
      resolve("");
      return;
    }
    const img = new Image();
    img.crossOrigin = "anonymous";
    img.onload = () => {
      const canvas = document.createElement("canvas");
      canvas.width = img.naturalWidth || 800;
      canvas.height = img.naturalHeight || 600;
      const ctx = canvas.getContext("2d");
      if (!ctx) {
        resolve("");
        return;
      }
      ctx.fillStyle = "#ffffff";
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
      resolve(canvas.toDataURL("image/png", 0.92));
    };
    img.onerror = () => resolve("");
    img.src = url;
  });

const renderPalletThumbnailDataUrl = (model, width = 300, height = 220) => {
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const rendererLocal = new THREE.WebGLRenderer({ canvas, antialias: true, alpha: true, preserveDrawingBuffer: true });
  rendererLocal.setPixelRatio(1);
  rendererLocal.setSize(width, height, false);

  const sceneLocal = new THREE.Scene();
  const cameraLocal = new THREE.PerspectiveCamera(34, width / height, 0.01, 20);
  cameraLocal.position.set(0, 0.95, 2.55);
  cameraLocal.lookAt(0, 0.52, 0);

  const key = new THREE.HemisphereLight("#ffffff", "#e8eef8", 1.06);
  const fill = new THREE.DirectionalLight("#ffffff", 0.72);
  fill.position.set(2, 3, 1.5);
  sceneLocal.add(key, fill);

  const group = new THREE.Group();
  sceneLocal.add(group);

  const stack = STACK_CONFIG[model.id] || { perLayer: Math.max(1, Math.floor(model.units / 2)), layers: 2 };
  const patternByPerLayer = {
    1: { cols: 1, rows: 1 },
    2: { cols: 2, rows: 1 },
    3: { cols: 3, rows: 1 },
    4: { cols: 2, rows: 2 },
  };
  const layout = patternByPerLayer[stack.perLayer] || { cols: Math.ceil(Math.sqrt(stack.perLayer)), rows: Math.ceil(stack.perLayer / Math.ceil(Math.sqrt(stack.perLayer))) };
  const cols = layout.cols;
  const rows = layout.rows;

  const modelL = model.size[0] / 1000;
  const modelW = model.size[1] / 1000;
  const modelH = model.size[2] / 1000;
  const footprintScale = 0.72;
  const baseL = modelL * footprintScale;
  const baseW = modelW * footprintScale;
  const totalH = modelH * footprintScale;

  const palletDeckH = 0.038;
  const palletFootH = 0.075;
  const cargoStartY = palletDeckH + palletFootH + 0.01;
  const layerH = Math.max(0.06, totalH / stack.layers);

  const marginL = baseL * 0.045;
  const marginW = baseW * 0.045;
  const gapL = baseL * 0.02;
  const gapW = baseW * 0.02;
  const itemL = (baseL - marginL * 2 - gapL * (cols - 1)) / cols;
  const itemW = (baseW - marginW * 2 - gapW * (rows - 1)) / rows;

  const cartonGeom = new THREE.BoxGeometry(itemL, layerH * 0.94, itemW);
  const cartonMat = new THREE.MeshStandardMaterial({ color: "#e5d2af", roughness: 0.78, metalness: 0.02 });
  const edgeMat = new THREE.LineBasicMaterial({ color: "#b89c6d" });
  for (let l = 0; l < stack.layers; l += 1) {
    for (let i = 0; i < stack.perLayer; i += 1) {
      const c = i % cols;
      const r = Math.floor(i / cols);
      const mesh = new THREE.Mesh(cartonGeom, cartonMat);
      mesh.position.set(
        -baseL / 2 + marginL + c * (itemL + gapL) + itemL / 2,
        cargoStartY + (l + 0.5) * layerH,
        -baseW / 2 + marginW + r * (itemW + gapW) + itemW / 2
      );
      mesh.add(new THREE.LineSegments(new THREE.EdgesGeometry(cartonGeom), edgeMat));
      group.add(mesh);
    }
  }

  const deckGeom = new THREE.BoxGeometry(baseL + 0.045, palletDeckH, baseW + 0.045);
  const deckMat = new THREE.MeshStandardMaterial({ color: "#d2b789", roughness: 0.82, metalness: 0.01 });
  const deck = new THREE.Mesh(deckGeom, deckMat);
  deck.position.y = palletFootH + palletDeckH / 2;
  group.add(deck);

  const runnerGeom = new THREE.BoxGeometry(baseL * 0.27, palletFootH, baseW * 0.14);
  const runnerMat = new THREE.MeshStandardMaterial({ color: "#c9aa75", roughness: 0.84, metalness: 0.01 });
  [
    [-baseL * 0.33, palletFootH / 2, -baseW * 0.30],
    [0, palletFootH / 2, -baseW * 0.30],
    [baseL * 0.33, palletFootH / 2, -baseW * 0.30],
    [-baseL * 0.33, palletFootH / 2, baseW * 0.30],
    [0, palletFootH / 2, baseW * 0.30],
    [baseL * 0.33, palletFootH / 2, baseW * 0.30],
  ].forEach(([x, y, z]) => {
    const foot = new THREE.Mesh(runnerGeom, runnerMat);
    foot.position.set(x, y, z);
    group.add(foot);
  });

  group.rotation.y = -0.36;
  rendererLocal.render(sceneLocal, cameraLocal);
  const dataUrl = canvas.toDataURL("image/png", 0.92);

  group.traverse((obj) => {
    if (obj.isMesh && obj.geometry) obj.geometry.dispose();
    if (obj.isMesh && obj.material) {
      if (Array.isArray(obj.material)) obj.material.forEach((m) => m.dispose?.());
      else obj.material.dispose?.();
    }
  });
  rendererLocal.dispose();
  return dataUrl;
};

const exportLoadingPlanPdf = async () => {
  const PdfCtor = window.jspdf?.jsPDF;
  if (!PdfCtor) {
    showPlacementHint("PDF engine failed to load. Please refresh and retry.", "warn");
    return;
  }
  if (!state.placed.length) {
    showPlacementHint("Load at least one pallet before exporting PDF.", "warn");
    return;
  }
  if (refs.exportPdf) refs.exportPdf.disabled = true;
  const data = buildPdfReportData();
  const doc = new PdfCtor({ orientation: "portrait", unit: "pt", format: "a4", compress: true });
  const pageW = doc.internal.pageSize.getWidth();
  const pageH = doc.internal.pageSize.getHeight();
  const margin = 30;
  let y = margin;

  const card = (x, top, w, h) => {
    doc.setFillColor(255, 255, 255);
    doc.setDrawColor(232, 236, 244);
    doc.roundedRect(x, top, w, h, 10, 10, "FD");
  };
  const addPageIfNeed = (need) => {
    if (y + need <= pageH - margin) return;
    doc.addPage();
    y = margin;
  };

  const containerShot = (() => {
    try {
      syncViewportAndScene();
      return refs.canvas.toDataURL("image/png", 0.9);
    } catch {
      return "";
    }
  })();

  const productImages = {};
  for (const row of data.selectedProducts) {
    productImages[row.product] = await loadImageDataUrl(row.imageUrl);
  }
  const palletThumbs = {};
  data.palletRows.forEach((row) => {
    const model = PALLET_MODELS.find((m) => m.id === row.modelId);
    if (model) {
      palletThumbs[row.modelId] = renderPalletThumbnailDataUrl(model);
    }
  });

  doc.setFillColor(255, 255, 255);
  doc.rect(0, 0, pageW, pageH, "F");

  card(margin, y, pageW - margin * 2, 88);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(16);
  doc.setTextColor(16, 24, 40);
  doc.text("Wattsonic Container Loading Plan", margin + 14, y + 28);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(10);
  doc.setTextColor(98, 110, 133);
  doc.text(`Container: ${data.container.label} (${data.containerType})  |  ${data.containerIndex}/${data.totalContainers}`, margin + 14, y + 48);
  doc.text(`Generated: ${data.generatedAt}`, margin + 14, y + 64);
  y += 104;

  const kpiW = (pageW - margin * 2 - 18) / 4;
  [
    ["Base Units", String(data.units)],
    ["Pallet Count", String(data.pallets)],
    ["Loading Rate", `${data.loadRate.toFixed(1)}%`],
    ["Weight Rate", `${data.weightRate.toFixed(1)}%`],
  ].forEach(([k, v], i) => {
    const x = margin + i * (kpiW + 6);
    card(x, y, kpiW, 58);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(9);
    doc.setTextColor(106, 116, 139);
    doc.text(k, x + 10, y + 20);
    doc.setFont("helvetica", "bold");
    doc.setFontSize(15);
    doc.setTextColor(17, 26, 43);
    doc.text(v, x + 10, y + 42);
  });
  y += 72;

  addPageIfNeed(228);
  card(margin, y, pageW - margin * 2, 214);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.setTextColor(17, 26, 43);
  doc.text("Container 3D Loading State", margin + 12, y + 20);
  if (containerShot) {
    doc.addImage(containerShot, "PNG", margin + 12, y + 28, pageW - margin * 2 - 24, 176, undefined, "FAST");
  }
  y += 228;

  addPageIfNeed(210);
  const prodCardH = 194;
  card(margin, y, pageW - margin * 2, prodCardH);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.setTextColor(17, 26, 43);
  doc.text("Selected Base Products (Image / Demand / Loaded)", margin + 12, y + 20);
  const colW = (pageW - margin * 2 - 36) / 3;
  const rows = data.selectedProducts.slice(0, 3);
  rows.forEach((item, idx) => {
    const x = margin + 12 + idx * (colW + 6);
    card(x, y + 28, colW, 154);
    const img = productImages[item.product];
    if (img) {
      doc.addImage(img, "PNG", x + 8, y + 36, colW - 16, 88, undefined, "FAST");
    }
    doc.setFont("helvetica", "bold");
    doc.setFontSize(8.5);
    doc.setTextColor(22, 31, 49);
    const title = doc.splitTextToSize(item.product, colW - 16);
    doc.text(title, x + 8, y + 132);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8);
    doc.setTextColor(95, 108, 131);
    doc.text(`Demand: ${item.demand} pcs`, x + 8, y + 147);
    doc.text(`Loaded: ${item.loadedUnits} pcs`, x + 8, y + 160);
  });
  y += prodCardH + 12;

  const maxPalletCards = Math.max(1, Math.min(6, data.palletRows.length));
  const palletCardH = 156;
  const cols = 2;
  const rowsCount = Math.ceil(maxPalletCards / cols);
  addPageIfNeed(rowsCount * (palletCardH + 10) + 34);
  card(margin, y, pageW - margin * 2, rowsCount * (palletCardH + 10) + 24);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.setTextColor(17, 26, 43);
  doc.text("Pallet Specs (3D / Size / Weight / Quantity)", margin + 12, y + 20);
  const gridW = pageW - margin * 2 - 24;
  const itemW = (gridW - 8) / 2;
  data.palletRows.slice(0, maxPalletCards).forEach((row, i) => {
    const col = i % cols;
    const r = Math.floor(i / cols);
    const x = margin + 12 + col * (itemW + 8);
    const top = y + 28 + r * (palletCardH + 10);
    card(x, top, itemW, palletCardH);
    const thumb = palletThumbs[row.modelId];
    if (thumb) {
      doc.addImage(thumb, "PNG", x + 8, top + 8, 108, 88, undefined, "FAST");
    }
    doc.setFont("helvetica", "bold");
    doc.setFontSize(8.6);
    doc.setTextColor(18, 27, 44);
    doc.text(`${row.product} [Type ${row.model}]`, x + 122, top + 20);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8);
    doc.setTextColor(96, 108, 130);
    doc.text(`Size: ${row.size[0]} x ${row.size[1]} x ${row.size[2]} mm`, x + 122, top + 38);
    doc.text(`Net Pallet Weight: ${Math.round(row.palletWeightKg)} kg`, x + 122, top + 52);
    doc.text(`Pallet Qty: ${row.pallets}`, x + 122, top + 66);
    doc.text(`Units per Pallet: ${row.unitsPerPallet} pcs`, x + 122, top + 80);
    doc.text(`Loaded Units: ${row.units} pcs`, x + 122, top + 94);
  });
  y += rowsCount * (palletCardH + 10) + 36;

  addPageIfNeed(96);
  card(margin, y, pageW - margin * 2, 82);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.setTextColor(18, 27, 44);
  doc.text("Container Parameters", margin + 12, y + 20);
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8.6);
  doc.setTextColor(92, 105, 129);
  doc.text(
    `Internal Size: ${data.container.lengthMm} x ${data.container.widthMm} x ${data.container.heightMm} mm   |   Max Payload: ${data.container.maxPayloadKg} kg   |   Max Gross: ${data.container.maxGrossKg} kg`,
    margin + 12,
    y + 40
  );
  doc.text(
    `Actual Payload: ${data.payloadKg} kg   |   Actual Gross: ${data.grossKg} kg`,
    margin + 12,
    y + 56
  );

  const safeType = data.container.label.replace(/[^a-z0-9]+/gi, "-").toLowerCase();
  const fileName = `wattsonic-loading-plan-${safeType}-${Date.now()}.pdf`;
  doc.save(fileName);
  showPlacementHint("Premium PDF downloaded.", "ok");
  if (refs.exportPdf) refs.exportPdf.disabled = false;
};

const exportLoadingPlanPdfSafe = () => {
  exportLoadingPlanPdf().catch((err) => {
    console.error("PDF export failed:", err);
    showPlacementHint("PDF export failed. Please retry.", "warn");
    if (refs.exportPdf) refs.exportPdf.disabled = false;
  });
};

const boot = async () => {
  await loadContainerCatalog();
  setupUI();
  init3D();
  updateStats();
  window.setTimeout(() => {
    openTour(false);
  }, 700);
};

boot();
