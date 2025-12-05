// ====== FLOWSIM MATERIAL STUDIO - CORE LOGIC ======

// ====== GLOBAL VARIABLES AND STATE MANAGEMENT ======

let appState = {
    userMode: 'engineer', // 'engineer', 'educator', 'presenter'
    currentMaterial: null,
    selectedMaterials: [],
    simulationActive: false,
    canvasMode: 'select', // 'select', 'draw', 'erase', 'measure'
    activeTool: null,
    canvasZoom: 1,
    canvasOffset: { x: 0, y: 0 },
    pptConnected: false,
    liveSession: null
};

let materialLibrary = [];
let flowSimulation = null;
let canvasManager = null;
let scannerAI = null;

// ====== DOM ELEMENTS ======

const loadingScreen = document.querySelector('.loading-screen');
const appContainer = document.querySelector('.container');
const modeSelect = document.querySelector('.mode-select');
const pptIntegrationBtn = document.querySelector('.btn-ppt-integration');
const startSimulationBtn = document.querySelector('.btn-start-simulation');
const stopSimulationBtn = document.querySelector('.btn-stop-simulation');
const exportPptBtn = document.querySelector('.btn-export-ppt');
const materialGrid = document.querySelector('.material-grid');
const cameraFeed = document.querySelector('.camera-feed');
const cameraCanvas = document.querySelector('.camera-canvas');
const mainCanvas = document.querySelector('#mainCanvas');
const propertySliders = document.querySelectorAll('.property-slider');
const aiTagsContainer = document.querySelector('.ai-tags');
const searchInput = document.querySelector('.search-box input');
const categoryButtons = document.querySelectorAll('.category-btn');

// ====== INITIALIZATION ======

document.addEventListener('DOMContentLoaded', async () => {
    console.log('🚀 FlowSim Material Studio v1.0 Initializing...');
    
    // Initialize with smooth loading animation
    await initializeApp();
    
    // Load default material library
    await loadMaterialLibrary();
    
    // Initialize canvas and tools
    initializeCanvas();
    
    // Initialize event listeners
    setupEventListeners();
    
    // Check for PowerPoint integration
    checkPPTIntegration();
    
    // Hide loading screen
    setTimeout(() => {
        loadingScreen.style.opacity = '0';
        setTimeout(() => {
            loadingScreen.style.display = 'none';
            appContainer.style.opacity = '1';
            console.log('✅ Application ready!');
        }, 300);
    }, 1500);
});

// ====== APP INITIALIZATION ======

async function initializeApp() {
    try {
        // Load configuration
        const config = await fetchConfig();
        
        // Initialize AI scanner if available
        if (config.aiEnabled) {
            scannerAI = new MaterialScannerAI(config.aiEndpoint);
            console.log('🤖 AI Scanner initialized');
        }
        
        // Initialize simulation engine
        flowSimulation = new FlowSimulationEngine();
        
        // Set user mode from localStorage or default
        const savedMode = localStorage.getItem('flowsim_user_mode');
        if (savedMode) {
            appState.userMode = savedMode;
            modeSelect.value = savedMode;
            updateUIMode(savedMode);
        }
        
    } catch (error) {
        console.error('Initialization error:', error);
        showNotification('⚠️ Warning: Some features may not be available', 'warning');
    }
}

// ====== MATERIAL LIBRARY MANAGEMENT ======

async function loadMaterialLibrary() {
    try {
        // In production, this would fetch from API
        // For demo, we'll use sample materials
        materialLibrary = [
            {
                id: 'mat_001',
                name: 'Steel Ball',
                category: 'metals',
                type: 'particle',
                thumbnail: 'assets/materials/steel_ball.jpg',
                properties: {
                    density: 7.85,
                    friction: 0.3,
                    elasticity: 0.8,
                    size: 'medium',
                    color: '#808080'
                },
                tags: ['metal', 'dense', 'round', 'industrial'],
                aiConfidence: 0.95
            },
            {
                id: 'mat_002',
                name: 'Plastic Pellet',
                category: 'plastics',
                type: 'particle',
                thumbnail: 'assets/materials/plastic_pellet.jpg',
                properties: {
                    density: 1.2,
                    friction: 0.4,
                    elasticity: 0.6,
                    size: 'small',
                    color: '#FF6B6B'
                },
                tags: ['plastic', 'light', 'colorful'],
                aiConfidence: 0.88
            },
            {
                id: 'mat_003',
                name: 'Sand',
                category: 'granular',
                type: 'bulk',
                thumbnail: 'assets/materials/sand.jpg',
                properties: {
                    density: 1.6,
                    friction: 0.7,
                    elasticity: 0.2,
                    size: 'fine',
                    color: '#D4A76A'
                },
                tags: ['granular', 'fine', 'natural', 'abrasive'],
                aiConfidence: 0.92
            },
            {
                id: 'mat_004',
                name: 'Wood Chip',
                category: 'organic',
                type: 'particle',
                thumbnail: 'assets/materials/wood_chip.jpg',
                properties: {
                    density: 0.6,
                    friction: 0.5,
                    elasticity: 0.4,
                    size: 'medium',
                    color: '#8B4513'
                },
                tags: ['wood', 'organic', 'light', 'biomass'],
                aiConfidence: 0.85
            },
            {
                id: 'mat_005',
                name: 'Grain',
                category: 'agricultural',
                type: 'granular',
                thumbnail: 'assets/materials/grain.jpg',
                properties: {
                    density: 0.75,
                    friction: 0.6,
                    elasticity: 0.3,
                    size: 'small',
                    color: '#F4D03F'
                },
                tags: ['grain', 'food', 'agricultural', 'flowable'],
                aiConfidence: 0.90
            }
        ];
        
        renderMaterialGrid(materialLibrary);
        updateStats();
        
    } catch (error) {
        console.error('Error loading material library:', error);
        materialGrid.innerHTML = `
            <div class="error-state">
                <i class="fas fa-exclamation-triangle"></i>
                <p>Failed to load material library</p>
                <button onclick="loadMaterialLibrary()" class="btn btn-secondary">Retry</button>
            </div>
        `;
    }
}

function renderMaterialGrid(materials) {
    materialGrid.innerHTML = '';
    
    materials.forEach(material => {
        const materialCard = document.createElement('div');
        materialCard.className = 'material-card';
        materialCard.dataset.id = material.id;
        
        materialCard.innerHTML = `
            <div class="material-thumbnail">
                <img src="${material.thumbnail}" alt="${material.name}" onerror="this.src='assets/materials/default.jpg'">
                ${material.aiConfidence > 0.9 ? '<span class="ai-badge">AI</span>' : ''}
            </div>
            <div class="material-info">
                <div class="material-name">${material.name}</div>
                <div class="material-tags">
                    ${material.tags.slice(0, 2).map(tag => `<span class="material-tag">${tag}</span>`).join('')}
                </div>
            </div>
        `;
        
        materialCard.addEventListener('click', () => selectMaterial(material));
        materialGrid.appendChild(materialCard);
    });
    
    // Add sample material (for demo)
    const sampleCard = document.createElement('div');
    sampleCard.className = 'material-card new';
    sampleCard.innerHTML = `
        <div class="material-thumbnail">
            <div class="add-new-material">
                <i class="fas fa-plus"></i>
                <p>Scan New</p>
            </div>
        </div>
        <div class="material-info">
            <div class="material-name">Add New Material</div>
            <div class="material-tags">
                <span class="material-tag">scan</span>
                <span class="material-tag">upload</span>
            </div>
        </div>
    `;
    
    sampleCard.addEventListener('click', () => openScanner());
    materialGrid.appendChild(sampleCard);
}

// ====== MATERIAL SELECTION AND PROPERTIES ======

function selectMaterial(material) {
    appState.currentMaterial = material;
    
    // Update UI
    document.querySelectorAll('.material-card').forEach(card => {
        card.classList.remove('selected');
        if (card.dataset.id === material.id) {
            card.classList.add('selected');
        }
    });
    
    // Update current material display
    updateCurrentMaterialDisplay();
    
    // Update property sliders
    updatePropertyControls(material.properties);
    
    // Update AI tags
    updateAITags(material);
    
    // Add to selected materials if not already
    if (!appState.selectedMaterials.some(m => m.id === material.id)) {
        appState.selectedMaterials.push(material);
        updateMaterialMixer();
    }
    
    showNotification(`Selected: ${material.name}`, 'success');
}

function updatePropertyControls(properties) {
    propertySliders.forEach(slider => {
        const property = slider.dataset.property;
        if (properties[property] !== undefined) {
            // Convert to percentage for slider
            const value = properties[property];
            const max = slider.max || 100;
            slider.value = value * max;
            
            // Update display
            const display = slider.parentElement.querySelector('.property-value');
            if (display) {
                display.textContent = value.toFixed(2);
            }
        }
    });
}

function updateAITags(material) {
    aiTagsContainer.innerHTML = '';
    
    material.tags.forEach(tag => {
        const tagElement = document.createElement('span');
        tagElement.className = 'ai-tag';
        tagElement.textContent = tag;
        aiTagsContainer.appendChild(tagElement);
    });
    
    // Add confidence indicator
    const confidence = document.createElement('div');
    confidence.className = 'ai-confidence';
    confidence.innerHTML = `
        <i class="fas fa-brain"></i>
        Confidence: ${(material.aiConfidence * 100).toFixed(0)}%
    `;
    aiTagsContainer.appendChild(confidence);
}

// ====== MATERIAL MIXER ======

function updateMaterialMixer() {
    const mixerSlots = document.querySelector('.mixer-slots');
    mixerSlots.innerHTML = '';
    
    // Add selected materials
    appState.selectedMaterials.forEach((material, index) => {
        const slot = document.createElement('div');
        slot.className = 'mixer-slot';
        slot.innerHTML = `
            <div class="slot-content">
                <img src="${material.thumbnail}" alt="${material.name}">
                <span>${material.name}</span>
            </div>
        `;
        
        slot.addEventListener('click', () => removeFromMixer(index));
        mixerSlots.appendChild(slot);
        
        // Add plus sign between materials
        if (index < appState.selectedMaterials.length - 1) {
            const plus = document.createElement('div');
            plus.className = 'mixer-plus';
            plus.innerHTML = '<i class="fas fa-plus"></i>';
            mixerSlots.appendChild(plus);
        }
    });
    
    // Add mix button if we have materials
    if (appState.selectedMaterials.length >= 2) {
        const mixBtn = document.createElement('button');
        mixBtn.className = 'mixer-btn';
        mixBtn.innerHTML = '<i class="fas fa-blender"></i>';
        mixBtn.title = 'Mix selected materials';
        mixBtn.addEventListener('click', mixMaterials);
        
        const mixContainer = document.createElement('div');
        mixContainer.className = 'mix-action';
        mixContainer.appendChild(mixBtn);
        mixerSlots.appendChild(mixContainer);
    }
}

function mixMaterials() {
    if (appState.selectedMaterials.length < 2) {
        showNotification('Please select at least 2 materials to mix', 'warning');
        return;
    }
    
    showNotification('Mixing materials...', 'info');
    
    // Calculate average properties
    const mixedProperties = {
        density: 0,
        friction: 0,
        elasticity: 0
    };
    
    appState.selectedMaterials.forEach(material => {
        mixedProperties.density += material.properties.density;
        mixedProperties.friction += material.properties.friction;
        mixedProperties.elasticity += material.properties.elasticity;
    });
    
    const count = appState.selectedMaterials.length;
    mixedProperties.density /= count;
    mixedProperties.friction /= count;
    mixedProperties.elasticity /= count;
    
    // Create mixed material
    const mixedMaterial = {
        id: 'mixed_' + Date.now(),
        name: 'Mixed Material',
        category: 'composite',
        type: 'custom',
        thumbnail: 'assets/materials/mixed.jpg',
        properties: mixedProperties,
        tags: ['mixed', 'composite', 'custom'],
        aiConfidence: 0.7,
        sourceMaterials: appState.selectedMaterials.map(m => m.id)
    };
    
    // Add to library and select
    materialLibrary.push(mixedMaterial);
    selectMaterial(mixedMaterial);
    
    showNotification('Materials mixed successfully!', 'success');
}

function removeFromMixer(index) {
    appState.selectedMaterials.splice(index, 1);
    updateMaterialMixer();
}

// ====== CANVAS MANAGEMENT ======

function initializeCanvas() {
    canvasManager = new CanvasManager('mainCanvas');
    
    // Setup tools
    const tools = document.querySelectorAll('.tool-btn');
    tools.forEach(tool => {
        tool.addEventListener('click', () => {
            const toolType = tool.dataset.tool;
            activateTool(toolType);
        });
    });
    
    // Setup background upload
    const backgroundUpload = document.getElementById('backgroundUpload');
    if (backgroundUpload) {
        backgroundUpload.addEventListener('change', handleBackgroundUpload);
    }
}

function activateTool(toolType) {
    appState.activeTool = toolType;
    appState.canvasMode = toolType;
    
    // Update UI
    document.querySelectorAll('.tool-btn').forEach(btn => {
        btn.classList.remove('active');
        if (btn.dataset.tool === toolType) {
            btn.classList.add('active');
        }
    });
    
    // Set canvas mode
    if (canvasManager) {
        canvasManager.setMode(toolType);
    }
    
    showNotification(`Tool activated: ${toolType}`, 'info');
}

function handleBackgroundUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = function(e) {
        const img = new Image();
        img.onload = function() {
            if (canvasManager) {
                canvasManager.setBackgroundImage(img);
                showNotification('Background image loaded', 'success');
            }
        };
        img.src = e.target.result;
    };
    reader.readAsDataURL(file);
}

// ====== SIMULATION ENGINE ======

class FlowSimulationEngine {
    constructor() {
        this.particles = [];
        this.isRunning = false;
        this.speed = 1.0;
        this.gravity = 9.8;
        this.friction = 0.98;
    }
    
    start() {
        if (this.isRunning) return;
        
        this.isRunning = true;
        appState.simulationActive = true;
        
        // Initialize particles based on current material
        if (appState.currentMaterial) {
            this.initializeParticles();
        }
        
        // Start animation loop
        this.animate();
        
        updateSimulationUI(true);
        showNotification('Simulation started', 'success');
    }
    
    stop() {
        this.isRunning = false;
        appState.simulationActive = false;
        
        updateSimulationUI(false);
        showNotification('Simulation stopped', 'info');
    }
    
    initializeParticles() {
        this.particles = [];
        const material = appState.currentMaterial;
        
        // Create particles based on material type
        let particleCount = 100;
        
        switch (material.type) {
            case 'bulk':
                particleCount = 500;
                break;
            case 'granular':
                particleCount = 300;
                break;
            case 'particle':
            default:
                particleCount = 100;
        }
        
        for (let i = 0; i < particleCount; i++) {
            this.particles.push({
                x: Math.random() * mainCanvas.width,
                y: Math.random() * 100,
                vx: (Math.random() - 0.5) * 2,
                vy: Math.random() * 2,
                radius: material.properties.size === 'small' ? 3 : 
                       material.properties.size === 'medium' ? 5 : 8,
                color: material.properties.color || '#3498db',
                density: material.properties.density,
                friction: material.properties.friction
            });
        }
    }
    
    animate() {
        if (!this.isRunning) return;
        
        // Update particle positions
        this.updateParticles();
        
        // Draw particles
        this.drawParticles();
        
        // Continue animation
        requestAnimationFrame(() => this.animate());
    }
    
    updateParticles() {
        const ctx = mainCanvas.getContext('2d');
        
        this.particles.forEach(particle => {
            // Apply gravity
            particle.vy += (this.gravity * particle.density) * 0.01;
            
            // Apply friction
            particle.vx *= particle.friction;
            particle.vy *= particle.friction;
            
            // Update position
            particle.x += particle.vx * this.speed;
            particle.y += particle.vy * this.speed;
            
            // Boundary collision
            if (particle.x < particle.radius || particle.x > mainCanvas.width - particle.radius) {
                particle.vx *= -0.8; // Bounce with energy loss
            }
            if (particle.y > mainCanvas.height - particle.radius) {
                particle.y = mainCanvas.height - particle.radius;
                particle.vy *= -0.8; // Bounce
            }
        });
    }
    
    drawParticles() {
        const ctx = mainCanvas.getContext('2d');
        
        // Clear canvas (with slight transparency for trail effect)
        ctx.fillStyle = 'rgba(255, 255, 255, 0.1)';
        ctx.fillRect(0, 0, mainCanvas.width, mainCanvas.height);
        
        // Draw each particle
        this.particles.forEach(particle => {
            ctx.beginPath();
            ctx.arc(particle.x, particle.y, particle.radius, 0, Math.PI * 2);
            ctx.fillStyle = particle.color;
            ctx.fill();
            ctx.closePath();
        });
    }
    
    setSpeed(speed) {
        this.speed = speed;
    }
    
    setGravity(gravity) {
        this.gravity = gravity;
    }
}

// ====== CANVAS MANAGER CLASS ======

class CanvasManager {
    constructor(canvasId) {
        this.canvas = document.getElementById(canvasId);
        this.ctx = this.canvas.getContext('2d');
        this.mode = 'select';
        this.backgroundImage = null;
        this.objects = [];
        this.isDrawing = false;
        this.lastPoint = null;
        
        this.setupEventListeners();
        this.resizeCanvas();
    }
    
    setupEventListeners() {
        // Canvas resize
        window.addEventListener('resize', () => this.resizeCanvas());
        
        // Drawing events
        this.canvas.addEventListener('mousedown', (e) => this.handleMouseDown(e));
        this.canvas.addEventListener('mousemove', (e) => this.handleMouseMove(e));
        this.canvas.addEventListener('mouseup', () => this.handleMouseUp());
        this.canvas.addEventListener('mouseleave', () => this.handleMouseUp());
        
        // Touch events for mobile
        this.canvas.addEventListener('touchstart', (e) => this.handleTouchStart(e));
        this.canvas.addEventListener('touchmove', (e) => this.handleTouchMove(e));
        this.canvas.addEventListener('touchend', () => this.handleMouseUp());
    }
    
    resizeCanvas() {
        const container = this.canvas.parentElement;
        this.canvas.width = container.clientWidth;
        this.canvas.height = container.clientHeight;
        this.draw();
    }
    
    setMode(mode) {
        this.mode = mode;
        this.updateCursor();
    }
    
    updateCursor() {
        switch (this.mode) {
            case 'draw':
                this.canvas.style.cursor = 'crosshair';
                break;
            case 'erase':
                this.canvas.style.cursor = 'url("assets/icons/eraser.png") 0 16, auto';
                break;
            case 'measure':
                this.canvas.style.cursor = 'ne-resize';
                break;
            default:
                this.canvas.style.cursor = 'default';
        }
    }
    
    setBackgroundImage(image) {
        this.backgroundImage = image;
        this.draw();
    }
    
    handleMouseDown(e) {
        const rect = this.canvas.getBoundingClientRect();
        const x = e.clientX - rect.left;
        const y = e.clientY - rect.top;
        
        this.isDrawing = true;
        this.lastPoint = { x, y };
        
        switch (this.mode) {
            case 'draw':
                this.startDrawing(x, y);
                break;
            case 'erase':
                this.eraseAt(x, y);
                break;
        }
    }
    
    handleMouseMove(e) {
        if (!this.isDrawing) return;
        
        const rect = this.canvas.getBoundingClientRect();
        const x = e.clientX - rect.left;
        const y = e.clientY - rect.top;
        
        switch (this.mode) {
            case 'draw':
                this.continueDrawing(x, y);
                break;
            case 'erase':
                this.eraseAt(x, y);
                break;
        }
        
        this.lastPoint = { x, y };
    }
    
    handleMouseUp() {
        this.isDrawing = false;
        this.lastPoint = null;
    }
    
    handleTouchStart(e) {
        e.preventDefault();
        const touch = e.touches[0];
        const mouseEvent = new MouseEvent('mousedown', {
            clientX: touch.clientX,
            clientY: touch.clientY
        });
        this.canvas.dispatchEvent(mouseEvent);
    }
    
    handleTouchMove(e) {
        e.preventDefault();
        const touch = e.touches[0];
        const mouseEvent = new MouseEvent('mousemove', {
            clientX: touch.clientX,
            clientY: touch.clientY
        });
        this.canvas.dispatchEvent(mouseEvent);
    }
    
    startDrawing(x, y) {
        this.objects.push({
            type: 'path',
            points: [{ x, y }],
            color: appState.currentMaterial?.properties.color || '#3498db',
            width: 3
        });
    }
    
    continueDrawing(x, y) {
        const currentObject = this.objects[this.objects.length - 1];
        if (currentObject && currentObject.type === 'path') {
            currentObject.points.push({ x, y });
            this.draw();
        }
    }
    
    eraseAt(x, y) {
        const eraseRadius = 20;
        
        this.objects = this.objects.filter(obj => {
            if (obj.type === 'path') {
                // Check if any point is within erase radius
                return !obj.points.some(point => {
                    const distance = Math.sqrt(
                        Math.pow(point.x - x, 2) + Math.pow(point.y - y, 2)
                    );
                    return distance < eraseRadius;
                });
            }
            return true;
        });
        
        this.draw();
    }
    
    draw() {
        this.ctx.clearRect(0, 0, this.canvas.width, this.canvas.height);
        
        // Draw background image if exists
        if (this.backgroundImage) {
            this.ctx.drawImage(this.backgroundImage, 0, 0, this.canvas.width, this.canvas.height);
        }
        
        // Draw objects
        this.objects.forEach(obj => {
            if (obj.type === 'path' && obj.points.length > 1) {
                this.ctx.beginPath();
                this.ctx.moveTo(obj.points[0].x, obj.points[0].y);
                
                for (let i = 1; i < obj.points.length; i++) {
                    this.ctx.lineTo(obj.points[i].x, obj.points[i].y);
                }
                
                this.ctx.strokeStyle = obj.color;
                this.ctx.lineWidth = obj.width;
                this.ctx.lineCap = 'round';
                this.ctx.lineJoin = 'round';
                this.ctx.stroke();
            }
        });
        
        // Draw current material if available
        this.drawCurrentMaterial();
    }
    
    drawCurrentMaterial() {
        if (!appState.currentMaterial || !appState.simulationActive) return;
        
        // This would be more sophisticated in a real implementation
        const material = appState.currentMaterial;
        this.ctx.fillStyle = material.properties.color;
        this.ctx.font = '14px Arial';
        this.ctx.fillText(material.name, 20, 30);
    }
    
    clear() {
        this.objects = [];
        this.draw();
    }
    
    exportAsImage() {
        return this.canvas.toDataURL('image/png');
    }
}

// ====== POWERPOINT INTEGRATION ======

function checkPPTIntegration() {
    // Check if PowerPoint is available (simulated)
    const isPPTConnected = localStorage.getItem('flowsim_ppt_connected') === 'true';
    appState.pptConnected = isPPTConnected;
    
    updatePPTStatus();
}

function connectToPowerPoint() {
    // Simulated PowerPoint connection
    showNotification('Connecting to PowerPoint...', 'info');
    
    setTimeout(() => {
        appState.pptConnected = true;
        localStorage.setItem('flowsim_ppt_connected', 'true');
        updatePPTStatus();
        showNotification('Successfully connected to PowerPoint!', 'success');
    }, 1500);
}

function updatePPTStatus() {
    const statusIndicator = document.querySelector('.status-indicator');
    const statusText = document.querySelector('.ppt-status p');
    
    if (appState.pptConnected) {
        statusIndicator.classList.add('connected');
        statusText.textContent = 'Connected to PowerPoint';
        pptIntegrationBtn.innerHTML = '<i class="fab fa-microsoft"></i> Export to PPT';
    } else {
        statusIndicator.classList.remove('connected');
        statusText.textContent = 'Not connected to PowerPoint';
        pptIntegrationBtn.innerHTML = '<i class="fab fa-microsoft"></i> Connect to PPT';
    }
}

async function exportToPowerPoint() {
    if (!appState.pptConnected) {
        connectToPowerPoint();
        return;
    }
    
    showNotification('Exporting to PowerPoint...', 'info');
    
    try {
        // Capture current canvas
        const canvasImage = canvasManager.exportAsImage();
        
        // Prepare simulation data
        const exportData = {
            material: appState.currentMaterial,
            simulationSettings: {
                speed: flowSimulation.speed,
                gravity: flowSimulation.gravity
            },
            timestamp: new Date().toISOString(),
            userMode: appState.userMode
        };
        
        // Generate VBA macro for automation
        const vbaCode = generateVBAMacro(exportData, canvasImage);
        
        // Show export options modal
        openExportModal(vbaCode, canvasImage);
        
    } catch (error) {
        console.error('Export error:', error);
        showNotification('Export failed. Please try again.', 'error');
    }
}

function generateVBAMacro(data, imageData) {
    // Generate VBA code for PowerPoint automation
    // This is a simplified example
    return `
' FlowSim Material Studio - PowerPoint Automation Macro
' Generated on ${new Date().toLocaleString()}
' Material: ${data.material.name}

Sub AddFlowSimSlide()
    Dim slide As slide
    Dim shp As Shape
    
    ' Create new slide
    Set slide = ActivePresentation.Slides.Add(ActivePresentation.Slides.Count + 1, ppLayoutBlank)
    
    ' Add title
    Set shp = slide.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 50)
    shp.TextFrame.TextRange.Text = "Flow Simulation: ${data.material.name}"
    shp.TextFrame.TextRange.Font.Size = 28
    
    ' Add material properties table
    ' ... (more VBA code) ...
End Sub

' ${data.material.name} Properties:
' - Density: ${data.material.properties.density}
' - Friction: ${data.material.properties.friction}
' - Elasticity: ${data.material.properties.elasticity}
`;
}

// ====== LIVE PRESENTATION MODE ======

function startLiveSession() {
    if (appState.liveSession) {
        showNotification('Live session already active', 'warning');
        return;
    }
    
    showNotification('Starting live presentation session...', 'info');
    
    // Generate session ID
    const sessionId = 'FS' + Date.now().toString(36).toUpperCase();
    
    appState.liveSession = {
        id: sessionId,
        startTime: new Date(),
        participants: 0,
        qrCode: generateQRCode(sessionId),
        controls: {
            allowVoting: true,
            allowQuestions: true,
            autoAdvance: false
        }
    };
    
    updateLiveSessionUI();
    showNotification(`Live session started: ${sessionId}`, 'success');
}

function generateQRCode(sessionId) {
    // In production, use a QR code library
    // For demo, return a placeholder
    return `https://flowsim.example.com/live/${sessionId}`;
}

function updateLiveSessionUI() {
    const sessionInfo = document.querySelector('.session-info');
    const qrContainer = document.querySelector('.qr-container');
    
    if (appState.liveSession) {
        sessionInfo.innerHTML = `
            <div class="info-item">
                <i class="fas fa-hashtag"></i>
                Session ID: <code>${appState.liveSession.id}</code>
            </div>
            <div class="info-item">
                <i class="fas fa-users"></i>
                Participants: ${appState.liveSession.participants}
            </div>
            <div class="info-item">
                <i class="fas fa-clock"></i>
                Duration: 00:00
            </div>
        `;
        
        qrContainer.innerHTML = `
            <div class="qr-placeholder">
                <i class="fas fa-qrcode"></i>
                <p>Scan to join</p>
            </div>
        `;
    }
}

// ====== EVENT LISTENERS SETUP ======

function setupEventListeners() {
    // User mode switching
    modeSelect.addEventListener('change', (e) => {
        const mode = e.target.value;
        appState.userMode = mode;
        localStorage.setItem('flowsim_user_mode', mode);
        updateUIMode(mode);
        showNotification(`Mode switched to: ${mode}`, 'info');
    });
    
    // PowerPoint integration
    pptIntegrationBtn.addEventListener('click', exportToPowerPoint);
    
    // Simulation controls
    startSimulationBtn?.addEventListener('click', () => {
        if (flowSimulation) {
            flowSimulation.start();
        }
    });
    
    stopSimulationBtn?.addEventListener('click', () => {
        if (flowSimulation) {
            flowSimulation.stop();
        }
    });
    
    // Property sliders
    propertySliders.forEach(slider => {
        slider.addEventListener('input', (e) => {
            const property = e.target.dataset.property;
            const value = parseFloat(e.target.value) / (e.target.max || 100);
            
            if (appState.currentMaterial) {
                appState.currentMaterial.properties[property] = value;
                
                // Update display
                const display = e.target.parentElement.querySelector('.property-value');
                if (display) {
                    display.textContent = value.toFixed(2);
                }
                
                // Update simulation if running
                if (flowSimulation && appState.simulationActive) {
                    // Update simulation parameters based on property
                    switch (property) {
                        case 'density':
                            // Adjust simulation based on density
                            break;
                        case 'friction':
                            if (flowSimulation) {
                                flowSimulation.friction = value;
                            }
                            break;
                    }
                }
            }
        });
    });
    
    // Search functionality
    searchInput?.addEventListener('input', (e) => {
        const query = e.target.value.toLowerCase();
        const filtered = materialLibrary.filter(material => 
            material.name.toLowerCase().includes(query) ||
            material.tags.some(tag => tag.toLowerCase().includes(query))
        );
        renderMaterialGrid(filtered);
    });
    
    // Category filtering
    categoryButtons?.forEach(btn => {
        btn.addEventListener('click', () => {
            const category = btn.dataset.category;
            
            // Update active state
            categoryButtons.forEach(b => b.classList.remove('active'));
            btn.classList.add('active');
            
            // Filter materials
            let filtered = materialLibrary;
            if (category !== 'all') {
                filtered = materialLibrary.filter(m => m.category === category);
            }
            renderMaterialGrid(filtered);
        });
    });
    
    // Live presentation controls
    const startLiveBtn = document.querySelector('.btn-start-live');
    startLiveBtn?.addEventListener('click', startLiveSession);
    
    // Export buttons
    exportPptBtn?.addEventListener('click', exportToPowerPoint);
    
    // Material mixer controls
    const clearMixerBtn = document.querySelector('.btn-clear-mixer');
    clearMixerBtn?.addEventListener('click', () => {
        appState.selectedMaterials = [];
        updateMaterialMixer();
        showNotification('Mixer cleared', 'info');
    });
}

// ====== UI UPDATES ======

function updateUIMode(mode) {
    // Update CSS variables based on mode
    document.documentElement.style.setProperty('--primary', getModeColor(mode));
    
    // Update UI elements
    const modeBadge = document.querySelector('.mode-badge');
    if (modeBadge) {
        modeBadge.textContent = mode.charAt(0).toUpperCase() + mode.slice(1);
        modeBadge.style.backgroundColor = getModeColor(mode);
    }
    
    // Show/hide mode-specific features
    updateModeSpecificFeatures(mode);
}

function getModeColor(mode) {
    switch (mode) {
        case 'educator': return '#2ecc71';
        case 'presenter': return '#9b59b6';
        case 'engineer':
        default: return '#3498db';
    }
}

function updateModeSpecificFeatures(mode) {
    // Show/hide features based on user mode
    const educatorFeatures = document.querySelectorAll('.educator-feature');
    const presenterFeatures = document.querySelectorAll('.presenter-feature');
    const engineerFeatures = document.querySelectorAll('.engineer-feature');
    
    switch (mode) {
        case 'educator':
            educatorFeatures.forEach(el => el.style.display = 'block');
            presenterFeatures.forEach(el => el.style.display = 'none');
            engineerFeatures.forEach(el => el.style.display = 'none');
            break;
        case 'presenter':
            educatorFeatures.forEach(el => el.style.display = 'none');
            presenterFeatures.forEach(el => el.style.display = 'block');
            engineerFeatures.forEach(el => el.style.display = 'none');
            break;
        case 'engineer':
        default:
            educatorFeatures.forEach(el => el.style.display = 'none');
            presenterFeatures.forEach(el => el.style.display = 'none');
            engineerFeatures.forEach(el => el.style.display = 'block');
    }
}

function updateCurrentMaterialDisplay() {
    const display = document.querySelector('.current-material-display');
    if (!display || !appState.currentMaterial) return;
    
    display.innerHTML = `
        <div class="material-thumbnail-small">
            <img src="${appState.currentMaterial.thumbnail}" 
                 alt="${appState.currentMaterial.name}"
                 onerror="this.src='assets/materials/default.jpg'">
        </div>
        <div class="material-info-small">
            <div class="material-name">${appState.currentMaterial.name}</div>
            <div class="material-properties">
                <span class="property-tag">ρ: ${appState.currentMaterial.properties.density.toFixed(2)}</span>
                <span class="property-tag">μ: ${appState.currentMaterial.properties.friction.toFixed(2)}</span>
            </div>
        </div>
        <button class="btn-tiny" onclick="clearCurrentMaterial()" title="Clear material">
            <i class="fas fa-times"></i>
        </button>
    `;
}

function clearCurrentMaterial() {
    appState.currentMaterial = null;
    updateCurrentMaterialDisplay();
    
    // Clear selected material cards
    document.querySelectorAll('.material-card.selected').forEach(card => {
        card.classList.remove('selected');
    });
    
    showNotification('Material cleared', 'info');
}

function updateSimulationUI(isRunning) {
    if (isRunning) {
        startSimulationBtn.style.display = 'none';
        stopSimulationBtn.style.display = 'block';
        
        // Update status bar
        updateStatusBar('Simulation running', 'success');
    } else {
        startSimulationBtn.style.display = 'block';
        stopSimulationBtn.style.display = 'none';
        
        updateStatusBar('Ready', 'info');
    }
}

function updateStatusBar(message, type) {
    const statusMessage = document.querySelector('.status-message');
    if (statusMessage) {
        statusMessage.innerHTML = `
            <i class="fas fa-${type === 'success' ? 'check-circle' : 'info-circle'}"></i>
            ${message}
        `;
        statusMessage.style.color = `var(--${type})`;
    }
}

function updateStats() {
    const totalMaterials = document.querySelector('.stat-total-materials');
    const popularMaterial = document.querySelector('.stat-popular-material');
    const simulationsRun = document.querySelector('.stat-simulations-run');
    
    if (totalMaterials) {
        totalMaterials.textContent = materialLibrary.length;
    }
    
    if (popularMaterial && materialLibrary.length > 0) {
        // Find most popular material (simplified)
        const popular = materialLibrary.reduce((prev, current) => 
            (prev.aiConfidence > current.aiConfidence) ? prev : current
        );
        popularMaterial.textContent = popular.name;
    }
    
    if (simulationsRun) {
        // Get from localStorage or default
        const runs = parseInt(localStorage.getItem('flowsim_simulations_run') || '0');
        simulationsRun.textContent = runs;
    }
}

// ====== UTILITY FUNCTIONS ======

async function fetchConfig() {
    // In production, fetch from server
    // For demo, return local config
    return {
        aiEnabled: true,
        aiEndpoint: 'https://api.flowsim.ai/v1/scan',
        maxFileSize: 10 * 1024 * 1024, // 10MB
        supportedFormats: ['image/jpeg', 'image/png', 'image/gif'],
        defaultMaterials: 5
    };
}

function showNotification(message, type = 'info') {
    // Create notification element
    const notification = document.createElement('div');
    notification.className = `notification notification-${type}`;
    notification.innerHTML = `
        <i class="fas fa-${getNotificationIcon(type)}"></i>
        <span>${message}</span>
        <button class="notification-close" onclick="this.parentElement.remove()">
            <i class="fas fa-times"></i>
        </button>
    `;
    
    // Add to notification container
    const container = document.querySelector('.notification-container') || createNotificationContainer();
    container.appendChild(notification);
    
    // Auto-remove after 5 seconds
    setTimeout(() => {
        if (notification.parentElement) {
            notification.remove();
        }
    }, 5000);
}

function getNotificationIcon(type) {
    switch (type) {
        case 'success': return 'check-circle';
        case 'warning': return 'exclamation-triangle';
        case 'error': return 'times-circle';
        default: return 'info-circle';
    }
}

function createNotificationContainer() {
    const container = document.createElement('div');
    container.className = 'notification-container';
    document.body.appendChild(container);
    return container;
}

function openScanner() {
    showNotification('Opening material scanner...', 'info');
    // This would open the scanner modal
    // For now, show a demo notification
    setTimeout(() => {
        showNotification('📸 Point camera at material to scan', 'info');
    }, 1000);
}

function openExportModal(vbaCode, canvasImage) {
    // Create modal for export options
    const modal = document.createElement('div');
    modal.className = 'modal visible';
    modal.innerHTML = `
        <div class="modal-content">
            <h3><i class="fas fa-file-export"></i> Export to PowerPoint</h3>
            
            <div class="export-options">
                <div class="export-grid">
                    <div class="export-option" data-type="slide">
                        <div class="export-icon">
                            <i class="fas fa-sliders-h"></i>
                        </div>
                        <div class="export-info">
                            <h5>Single Slide</h5>
                            <p>Current simulation as one slide</p>
                        </div>
                        <button class="export-btn" onclick="exportAsSlide()">
                            <i class="fas fa-arrow-right"></i>
                        </button>
                    </div>
                    
                    <div class="export-option" data-type="report">
                        <div class="export-icon">
                            <i class="fas fa-file-alt"></i>
                        </div>
                        <div class="export-info">
                            <h5>Full Report</h5>
                            <p>Multiple slides with analysis</p>
                        </div>
                        <button class="export-btn" onclick="exportAsReport()">
                            <i class="fas fa-arrow-right"></i>
                        </button>
                    </div>
                </div>
                
                <div class="export-settings">
                    <h5>Export Settings</h5>
                    <div class="settings-grid">
                        <div class="setting-item">
                            <label>Slide Layout</label>
                            <select id="slideLayout">
                                <option value="title">Title Slide</option>
                                <option value="content">Content Slide</option>
                                <option value="comparison">Comparison</option>
                            </select>
                        </div>
                        <div class="setting-item">
                            <label>Include VBA Code</label>
                            <select id="includeVBA">
                                <option value="yes">Yes</option>
                                <option value="no">No</option>
                            </select>
                        </div>
                    </div>
                </div>
                
                <div class="vba-generator">
                    <h5>VBA Macro Code</h5>
                    <textarea readonly>${vbaCode}</textarea>
                    <div class="vba-actions">
                        <button class="btn btn-secondary" onclick="copyVBACode()">
                            <i class="fas fa-copy"></i> Copy VBA
                        </button>
                        <button class="btn btn-primary" onclick="downloadVBACode()">
                            <i class="fas fa-download"></i> Download .bas
                        </button>
                    </div>
                </div>
            </div>
            
            <div class="modal-actions">
                <button class="btn btn-secondary" onclick="this.closest('.modal').remove()">
                    Cancel
                </button>
                <button class="btn btn-primary" onclick="confirmExport()">
                    <i class="fas fa-file-powerpoint"></i> Export Now
                </button>
            </div>
        </div>
    `;
    
    document.body.appendChild(modal);
}

// ====== EXPORT FUNCTIONS ======

function exportAsSlide() {
    showNotification('Exporting as single slide...', 'info');
    // Implementation would go here
}

function exportAsReport() {
    showNotification('Exporting as full report...', 'info');
    // Implementation would go here
}

function copyVBACode() {
    const textarea = document.querySelector('.vba-generator textarea');
    textarea.select();
    document.execCommand('copy');
    showNotification('VBA code copied to clipboard!', 'success');
}

function downloadVBACode() {
    const textarea = document.querySelector('.vba-generator textarea');
    const blob = new Blob([textarea.value], { type: 'text/plain' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `flowsim_macro_${Date.now()}.bas`;
    a.click();
    URL.revokeObjectURL(url);
    showNotification('VBA macro downloaded!', 'success');
}

function confirmExport() {
    showNotification('Exporting to PowerPoint...', 'info');
    
    // Simulate export process
    setTimeout(() => {
        showNotification('✅ Successfully exported to PowerPoint!', 'success');
        
        // Update stats
        const currentRuns = parseInt(localStorage.getItem('flowsim_simulations_run') || '0');
        localStorage.setItem('flowsim_simulations_run', (currentRuns + 1).toString());
        updateStats();
        
        // Close modal
        document.querySelector('.modal.visible')?.remove();
    }, 2000);
}

// ====== INITIALIZE ON LOAD ======

// Add notification styles to head
const style = document.createElement('style');
style.textContent = `
    .notification-container {
        position: fixed;
        top: 20px;
        right: 20px;
        z-index: 9999;
        max-width: 350px;
    }
    
    .notification {
        background: white;
        border-radius: 8px;
        padding: 12px 16px;
        margin-bottom: 10px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        display: flex;
        align-items: center;
        gap: 12px;
        animation: slideIn 0.3s ease;
        border-left: 4px solid var(--info);
    }
    
    .notification-success {
        border-left-color: var(--success);
    }
    
    .notification-warning {
        border-left-color: var(--warning);
    }
    
    .notification-error {
        border-left-color: var(--danger);
    }
    
    .notification i {
        font-size: 18px;
    }
    
    .notification-close {
        margin-left: auto;
        background: none;
        border: none;
        color: var(--gray);
        cursor: pointer;
        padding: 4px;
    }
    
    @keyframes slideIn {
        from {
            transform: translateX(100%);
            opacity: 0;
        }
        to {
            transform: translateX(0);
            opacity: 1;
        }
    }
`;
document.head.appendChild(style);

console.log('📦 FlowSim Material Studio - Core script loaded');
