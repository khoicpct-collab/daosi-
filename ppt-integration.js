// ====== FLOWSIM MATERIAL STUDIO - POWERPOINT INTEGRATION ======

class PowerPointIntegration {
    constructor() {
        this.officeJSLoaded = false;
        this.isConnected = false;
        this.currentPresentation = null;
        this.exportTemplates = {};
        this.vbaMacros = {};
        this.liveControl = null;
        
        // Export settings
        this.exportSettings = {
            slideLayout: 'content',
            includeVBA: true,
            includeAnimations: true,
            dataTables: true,
            simulationFrames: 5,
            exportFormat: 'pptx',
            theme: 'flowsim',
            autoOpen: true
        };
        
        // VBA macro templates
        this.vbaTemplates = {
            basic: this.getBasicVBATemplate(),
            advanced: this.getAdvancedVBATemplate(),
            simulation: this.getSimulationVBATemplate(),
            report: this.getReportVBATemplate()
        };
        
        this.init();
    }
    
    // ====== INITIALIZATION ======
    
    async init() {
        try {
            // Check if Office.js is available
            if (typeof Office !== 'undefined') {
                await this.loadOfficeJS();
                this.officeJSLoaded = true;
                console.log('✅ Office.js loaded');
            } else {
                console.warn('Office.js not available, using fallback export');
            }
            
            // Load export templates
            await this.loadTemplates();
            
            // Setup event listeners
            this.setupEventListeners();
            
            // Check for saved settings
            this.loadSettings();
            
            console.log('📊 PowerPoint Integration initialized');
            
        } catch (error) {
            console.error('PowerPoint integration init failed:', error);
        }
    }
    
    async loadOfficeJS() {
        // Office.js should be loaded via CDN in HTML
        // This is just a placeholder for initialization logic
        return new Promise((resolve) => {
            if (window.Office && Office.context) {
                this.isConnected = true;
                resolve();
            } else {
                // Try to initialize
                Office.onReady(() => {
                    this.isConnected = true;
                    resolve();
                });
            }
        });
    }
    
    async loadTemplates() {
        try {
            // Load templates from server or local storage
            const templates = await this.fetchTemplates();
            this.exportTemplates = templates;
            
            console.log(`📁 Loaded ${Object.keys(templates).length} templates`);
            
        } catch (error) {
            console.error('Failed to load templates:', error);
            // Load default templates
            this.exportTemplates = this.getDefaultTemplates();
        }
    }
    
    // ====== CONNECTION MANAGEMENT ======
    
    async connectToPowerPoint() {
        if (this.isConnected) {
            return true;
        }
        
        showNotification('Connecting to PowerPoint...', 'info');
        
        try {
            if (this.officeJSLoaded) {
                // Use Office.js API
                await this.connectWithOfficeJS();
            } else {
                // Fallback: Simulate connection
                await this.simulateConnection();
            }
            
            this.isConnected = true;
            this.updateConnectionStatus();
            showNotification('✅ Connected to PowerPoint!', 'success');
            
            return true;
            
        } catch (error) {
            console.error('Connection failed:', error);
            showNotification('❌ Failed to connect to PowerPoint', 'error');
            return false;
        }
    }
    
    async connectWithOfficeJS() {
        return new Promise((resolve, reject) => {
            Office.context.document.getFileAsync(Office.FileType.Compressed, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    this.currentPresentation = result.value;
                    resolve();
                } else {
                    reject(new Error('Could not access presentation'));
                }
            });
        });
    }
    
    async simulateConnection() {
        // Simulate connection delay
        return new Promise(resolve => {
            setTimeout(() => {
                this.currentPresentation = {
                    id: 'simulated_' + Date.now(),
                    name: 'FlowSim Presentation',
                    slides: []
                };
                resolve();
            }, 1500);
        });
    }
    
    disconnect() {
        this.isConnected = false;
        this.currentPresentation = null;
        this.updateConnectionStatus();
        showNotification('Disconnected from PowerPoint', 'info');
    }
    
    updateConnectionStatus() {
        const statusIndicator = document.querySelector('.status-indicator');
        const statusText = document.querySelector('.ppt-status p');
        const connectBtn = document.querySelector('.btn-ppt-integration');
        
        if (this.isConnected) {
            statusIndicator?.classList.add('connected');
            statusText && (statusText.textContent = 'Connected to PowerPoint');
            connectBtn && (connectBtn.innerHTML = '<i class="fab fa-microsoft"></i> Export to PPT');
            
            // Update stats
            this.updateExportStats();
            
        } else {
            statusIndicator?.classList.remove('connected');
            statusText && (statusText.textContent = 'Not connected to PowerPoint');
            connectBtn && (connectBtn.innerHTML = '<i class="fab fa-microsoft"></i> Connect to PPT');
        }
    }
    
    // ====== EXPORT FUNCTIONS ======
    
    async exportToPowerPoint(options = {}) {
        // Merge options with default settings
        const exportOptions = { ...this.exportSettings, ...options };
        
        // Check connection
        if (!this.isConnected) {
            const connected = await this.connectToPowerPoint();
            if (!connected) return false;
        }
        
        showNotification('Preparing PowerPoint export...', 'info');
        
        try {
            // Gather data for export
            const exportData = await this.prepareExportData(exportOptions);
            
            // Create presentation
            const presentation = await this.createPresentation(exportData, exportOptions);
            
            // Add slides based on options
            await this.addSlides(presentation, exportData, exportOptions);
            
            // Generate VBA macro if requested
            if (exportOptions.includeVBA) {
                await this.addVBAMacro(presentation, exportData);
            }
            
            // Save/export the presentation
            const result = await this.finalizeExport(presentation, exportOptions);
            
            // Show success message
            this.showExportSuccess(result, exportOptions);
            
            return result;
            
        } catch (error) {
            console.error('Export failed:', error);
            showNotification('Export failed: ' + error.message, 'error');
            return false;
        }
    }
    
    async prepareExportData(options) {
        const appState = window.appState || {};
        const simulation = window.flowSimulation;
        const canvasManager = window.canvasManager;
        
        // Capture current simulation state
        const simulationData = simulation ? this.captureSimulationData(simulation) : null;
        const canvasImage = canvasManager ? canvasManager.exportAsImage() : null;
        const materialData = appState.currentMaterial || {};
        const libraryData = window.materialLibrary || [];
        
        return {
            metadata: {
                exportDate: new Date().toISOString(),
                version: '1.0',
                userMode: appState.userMode || 'engineer'
            },
            simulation: simulationData,
            materials: {
                current: materialData,
                library: libraryData.slice(0, 10), // First 10 materials
                selected: appState.selectedMaterials || []
            },
            canvas: {
                image: canvasImage,
                objects: canvasManager?.objects || [],
                background: canvasManager?.backgroundImage?.src || null
            },
            analysis: this.generateAnalysisReport(materialData, simulationData),
            settings: options
        };
    }
    
    captureSimulationData(simulation) {
        return {
            isRunning: simulation.isRunning,
            speed: simulation.speed,
            gravity: simulation.gravity,
            friction: simulation.friction,
            particleCount: simulation.particles?.length || 0,
            particles: simulation.particles?.slice(0, 100) || [], // Sample particles
            timestamp: Date.now()
        };
    }
    
    generateAnalysisReport(material, simulation) {
        return {
            materialProperties: material.properties || {},
            flowCharacteristics: this.calculateFlowCharacteristics(simulation),
            recommendations: this.generateRecommendations(material, simulation),
            stats: {
                density: material.properties?.density || 0,
                friction: material.properties?.friction || 0,
                elasticity: material.properties?.elasticity || 0,
                flowRate: this.calculateFlowRate(simulation)
            }
        };
    }
    
    calculateFlowCharacteristics(simulation) {
        if (!simulation || !simulation.particles) {
            return { flowability: 'unknown', stability: 'unknown', velocity: 0 };
        }
        
        const particles = simulation.particles;
        let totalVelocity = 0;
        let velocityVariance = 0;
        
        particles.forEach(p => {
            const velocity = Math.sqrt(p.vx * p.vx + p.vy * p.vy);
            totalVelocity += velocity;
        });
        
        const avgVelocity = totalVelocity / particles.length;
        
        // Calculate variance
        particles.forEach(p => {
            const velocity = Math.sqrt(p.vx * p.vx + p.vy * p.vy);
            velocityVariance += Math.pow(velocity - avgVelocity, 2);
        });
        
        const stdDev = Math.sqrt(velocityVariance / particles.length);
        
        return {
            flowability: avgVelocity > 2 ? 'high' : avgVelocity > 1 ? 'medium' : 'low',
            stability: stdDev < 0.5 ? 'stable' : stdDev < 1 ? 'moderate' : 'unstable',
            averageVelocity: avgVelocity,
            velocityStdDev: stdDev,
            particleDistribution: this.analyzeParticleDistribution(particles)
        };
    }
    
    analyzeParticleDistribution(particles) {
        if (!particles.length) return 'uniform';
        
        const zones = {
            top: 0, middle: 0, bottom: 0,
            left: 0, center: 0, right: 0
        };
        
        particles.forEach(p => {
            // Vertical zones
            if (p.y < 0.33) zones.top++;
            else if (p.y < 0.66) zones.middle++;
            else zones.bottom++;
            
            // Horizontal zones
            if (p.x < 0.33) zones.left++;
            else if (p.x < 0.66) zones.center++;
            else zones.right++;
        });
        
        const verticalRatio = Math.max(zones.top, zones.middle, zones.bottom) / particles.length;
        const horizontalRatio = Math.max(zones.left, zones.center, zones.right) / particles.length;
        
        if (verticalRatio > 0.6 || horizontalRatio > 0.6) return 'clustered';
        if (verticalRatio < 0.4 && horizontalRatio < 0.4) return 'dispersed';
        return 'uniform';
    }
    
    calculateFlowRate(simulation) {
        if (!simulation || !simulation.particles) return 0;
        
        let totalFlow = 0;
        simulation.particles.forEach(p => {
            // Simple flow calculation based on velocity and density
            const velocity = Math.sqrt(p.vx * p.vx + p.vy * p.vy);
            const density = p.density || 1;
            totalFlow += velocity * density;
        });
        
        return totalFlow / simulation.particles.length;
    }
    
    generateRecommendations(material, simulation) {
        const recommendations = [];
        
        if (material.properties) {
            // Density recommendations
            if (material.properties.density > 5) {
                recommendations.push('High density material - consider heavier equipment');
            } else if (material.properties.density < 1) {
                recommendations.push('Low density material - may require containment');
            }
            
            // Friction recommendations
            if (material.properties.friction > 0.6) {
                recommendations.push('High friction - ensure proper surface treatment');
            } else if (material.properties.friction < 0.3) {
                recommendations.push('Low friction - consider anti-slip surfaces');
            }
            
            // Flow characteristics recommendations
            const flow = this.calculateFlowCharacteristics(simulation);
            if (flow.flowability === 'low') {
                recommendations.push('Low flowability - consider vibration or agitation');
            } else if (flow.flowability === 'high') {
                recommendations.push('High flowability - ensure proper containment');
            }
        }
        
        return recommendations.slice(0, 3); // Top 3 recommendations
    }
    
    // ====== PRESENTATION CREATION ======
    
    async createPresentation(data, options) {
        if (this.officeJSLoaded && this.currentPresentation) {
            // Use Office.js to create presentation
            return await this.createPresentationOfficeJS(data, options);
        } else {
            // Create simulated presentation
            return this.createPresentationSimulated(data, options);
        }
    }
    
    async createPresentationOfficeJS(data, options) {
        return new Promise((resolve, reject) => {
            try {
                const presentation = {
                    id: this.currentPresentation.id,
                    name: `FlowSim_${new Date().toISOString().slice(0, 10)}.pptx`,
                    slides: [],
                    theme: options.theme,
                    author: 'FlowSim Material Studio',
                    created: new Date().toISOString()
                };
                
                resolve(presentation);
            } catch (error) {
                reject(error);
            }
        });
    }
    
    createPresentationSimulated(data, options) {
        return {
            id: 'sim_pres_' + Date.now(),
            name: `FlowSim_Export_${Date.now()}.pptx`,
            slides: [],
            theme: options.theme,
            author: 'FlowSim Material Studio',
            created: new Date().toISOString(),
            fileSize: '0 KB',
            slideCount: 0
        };
    }
    
    async addSlides(presentation, data, options) {
        const slides = [];
        const slideLayout = options.slideLayout;
        
        // 1. Title slide
        slides.push(await this.createTitleSlide(data));
        
        // 2. Material properties slide
        slides.push(await this.createMaterialSlide(data));
        
        // 3. Simulation overview slide
        if (data.simulation) {
            slides.push(await this.createSimulationSlide(data));
        }
        
        // 4. Analysis and results slide
        slides.push(await this.createAnalysisSlide(data));
        
        // 5. Recommendations slide
        if (data.analysis.recommendations.length > 0) {
            slides.push(await this.createRecommendationsSlide(data));
        }
        
        // 6. Data table slide (if enabled)
        if (options.dataTables) {
            slides.push(await this.createDataTableSlide(data));
        }
        
        // 7. Export details slide
        slides.push(await this.createExportDetailsSlide(data));
        
        // Add all slides to presentation
        presentation.slides = slides;
        presentation.slideCount = slides.length;
        
        return slides;
    }
    
    async createTitleSlide(data) {
        const materialName = data.materials.current?.name || 'Current Material';
        
        return {
            id: 'slide_title_' + Date.now(),
            type: 'title',
            layout: 'title',
            content: {
                title: `Flow Simulation: ${materialName}`,
                subtitle: 'FlowSim Material Studio Analysis Report',
                date: new Date().toLocaleDateString(),
                author: data.metadata.userMode === 'educator' ? 'Educational Analysis' : 
                       data.metadata.userMode === 'presenter' ? 'Presentation Deck' : 'Engineering Report',
                logo: true
            },
            theme: 'flowsim_title',
            notes: 'Generated by FlowSim Material Studio - Professional Material Flow Simulation'
        };
    }
    
    async createMaterialSlide(data) {
        const material = data.materials.current;
        
        return {
            id: 'slide_material_' + Date.now(),
            type: 'content',
            layout: 'twoColumn',
            content: {
                title: 'Material Properties',
                leftColumn: {
                    type: 'properties',
                    data: [
                        { label: 'Name', value: material?.name || 'N/A' },
                        { label: 'Category', value: material?.category || 'N/A' },
                        { label: 'Type', value: material?.type || 'N/A' },
                        { label: 'Density', value: material?.properties?.density?.toFixed(2) || 'N/A' },
                        { label: 'Friction', value: material?.properties?.friction?.toFixed(2) || 'N/A' },
                        { label: 'Elasticity', value: material?.properties?.elasticity?.toFixed(2) || 'N/A' }
                    ]
                },
                rightColumn: {
                    type: 'visual',
                    content: material?.thumbnail ? {
                        type: 'image',
                        src: material.thumbnail,
                        caption: 'Material Sample'
                    } : {
                        type: 'colorBox',
                        color: material?.properties?.color || '#3498db',
                        caption: 'Material Color'
                    }
                }
            },
            theme: 'flowsim_content'
        };
    }
    
    async createSimulationSlide(data) {
        const simulation = data.simulation;
        const flow = data.analysis.flowCharacteristics;
        
        return {
            id: 'slide_simulation_' + Date.now(),
            type: 'content',
            layout: 'full',
            content: {
                title: 'Simulation Analysis',
                sections: [
                    {
                        title: 'Flow Characteristics',
                        type: 'stats',
                        data: [
                            { label: 'Flowability', value: flow.flowability, icon: 'fas fa-tachometer-alt' },
                            { label: 'Stability', value: flow.stability, icon: 'fas fa-balance-scale' },
                            { label: 'Avg Velocity', value: flow.averageVelocity?.toFixed(2) + ' u/s', icon: 'fas fa-running' },
                            { label: 'Distribution', value: flow.particleDistribution, icon: 'fas fa-layer-group' }
                        ]
                    },
                    {
                        title: 'Simulation Parameters',
                        type: 'parameters',
                        data: [
                            { label: 'Speed', value: simulation.speed?.toFixed(2) || 'N/A' },
                            { label: 'Gravity', value: simulation.gravity?.toFixed(2) || 'N/A' },
                            { label: 'Friction', value: simulation.friction?.toFixed(2) || 'N/A' },
                            { label: 'Particles', value: simulation.particleCount || 0 }
                        ]
                    }
                ]
            },
            theme: 'flowsim_simulation'
        };
    }
    
    async createAnalysisSlide(data) {
        const analysis = data.analysis;
        
        return {
            id: 'slide_analysis_' + Date.now(),
            type: 'content',
            layout: 'analysis',
            content: {
                title: 'Technical Analysis',
                charts: [
                    {
                        type: 'bar',
                        title: 'Material Properties',
                        data: [
                            { label: 'Density', value: analysis.stats.density, max: 10 },
                            { label: 'Friction', value: analysis.stats.friction, max: 1 },
                            { label: 'Elasticity', value: analysis.stats.elasticity, max: 1 }
                        ],
                        colors: ['#3498db', '#2ecc71', '#9b59b6']
                    },
                    {
                        type: 'gauge',
                        title: 'Flow Rate',
                        value: analysis.stats.flowRate,
                        max: 10,
                        segments: [
                            { from: 0, to: 3, color: '#e74c3c', label: 'Low' },
                            { from: 3, to: 7, color: '#f39c12', label: 'Medium' },
                            { from: 7, to: 10, color: '#2ecc71', label: 'High' }
                        ]
                    }
                ],
                summary: 'Flow analysis indicates ' + 
                        (analysis.stats.flowRate > 7 ? 'excellent flow characteristics' :
                         analysis.stats.flowRate > 4 ? 'adequate flow performance' :
                         'potential flow issues requiring attention')
            },
            theme: 'flowsim_analysis'
        };
    }
    
    async createRecommendationsSlide(data) {
        const recommendations = data.analysis.recommendations;
        
        return {
            id: 'slide_recommendations_' + Date.now(),
            type: 'content',
            layout: 'recommendations',
            content: {
                title: 'Recommendations & Best Practices',
                recommendations: recommendations.map((rec, index) => ({
                    id: index + 1,
                    text: rec,
                    icon: this.getRecommendationIcon(rec),
                    priority: this.getRecommendationPriority(rec)
                })),
                summary: 'Based on the simulation analysis, ' + 
                        (recommendations.length > 2 ? 'multiple improvements are recommended' :
                         recommendations.length > 0 ? 'some adjustments are suggested' :
                         'current setup appears optimal')
            },
            theme: 'flowsim_recommendations'
        };
    }
    
    async createDataTableSlide(data) {
        const material = data.materials.current;
        const library = data.materials.library;
        
        return {
            id: 'slide_data_' + Date.now(),
            type: 'content',
            layout: 'table',
            content: {
                title: 'Material Comparison',
                table: {
                    headers: ['Material', 'Category', 'Density', 'Friction', 'Elasticity', 'Flow Rate'],
                    rows: library.slice(0, 8).map(mat => [
                        mat.name,
                        mat.category,
                        mat.properties?.density?.toFixed(2) || 'N/A',
                        mat.properties?.friction?.toFixed(2) || 'N/A',
                        mat.properties?.elasticity?.toFixed(2) || 'N/A',
                        this.calculateMaterialFlowRate(mat)?.toFixed(2) || 'N/A'
                    ]),
                    highlightRow: library.findIndex(mat => mat.id === material?.id)
                },
                notes: 'Table shows comparison with other materials in library'
            },
            theme: 'flowsim_table'
        };
    }
    
    async createExportDetailsSlide(data) {
        return {
            id: 'slide_export_' + Date.now(),
            type: 'content',
            layout: 'details',
            content: {
                title: 'Export Details',
                details: [
                    { label: 'Export Date', value: new Date(data.metadata.exportDate).toLocaleString() },
                    { label: 'Report Type', value: this.getReportType(data.metadata.userMode) },
                    { label: 'Software Version', value: data.metadata.version },
                    { label: 'Total Slides', value: '7' },
                    { label: 'Analysis Confidence', value: 'High' },
                    { label: 'Export Format', value: 'PowerPoint (.pptx)' }
                ],
                footer: 'This report was automatically generated by FlowSim Material Studio. ' +
                       'For detailed analysis, please refer to the full simulation data.'
            },
            theme: 'flowsim_details'
        };
    }
    
    // ====== VBA MACRO GENERATION ======
    
    async addVBAMacro(presentation, data) {
        const vbaCode = this.generateVBAMacro(data);
        this.vbaMacros[presentation.id] = vbaCode;
        
        // Add VBA module to presentation
        presentation.vbaModule = {
            name: 'FlowSim_AutoUpdate',
            code: vbaCode,
            procedures: [
                'UpdateSimulationData',
                'RefreshCharts',
                'ExportToExcel',
                'GenerateReport'
            ]
        };
        
        return vbaCode;
    }
    
    generateVBAMacro(data) {
        const material = data.materials.current;
        const timestamp = new Date().toISOString();
        
        // Select template based on user mode
        let template = this.vbaTemplates.basic;
        
        if (data.metadata.userMode === 'engineer') {
            template = this.vbaTemplates.advanced;
        } else if (data.metadata.userMode === 'educator') {
            template = this.vbaTemplates.simulation;
        } else if (data.metadata.userMode === 'presenter') {
            template = this.vbaTemplates.report;
        }
        
        // Replace placeholders
        return template
            .replace(/{MATERIAL_NAME}/g, material?.name || 'Unknown Material')
            .replace(/{EXPORT_DATE}/g, timestamp)
            .replace(/{DENSITY}/g, material?.properties?.density?.toFixed(2) || '0.00')
            .replace(/{FRICTION}/g, material?.properties?.friction?.toFixed(2) || '0.00')
            .replace(/{ELASTICITY}/g, material?.properties?.elasticity?.toFixed(2) || '0.00')
            .replace(/{USER_MODE}/g, data.metadata.userMode || 'engineer')
            .replace(/{SLIDE_COUNT}/g, '7')
            .replace(/{FLOW_RATE}/g, data.analysis.stats.flowRate?.toFixed(2) || '0.00');
    }
    
    getBasicVBATemplate() {
        return `' FlowSim Material Studio - PowerPoint Automation Macro
' Generated on {EXPORT_DATE}
' Material: {MATERIAL_NAME}

Option Explicit

' Main procedure to update all slides
Sub UpdateFlowSimData()
    On Error Resume Next
    
    Dim pres As Presentation
    Dim slide As Slide
    Dim shp As Shape
    Dim i As Integer
    
    Set pres = ActivePresentation
    
    ' Update title slide
    Set slide = pres.Slides(1)
    For Each shp In slide.Shapes
        If shp.HasTextFrame Then
            If InStr(shp.TextFrame.TextRange.Text, "Flow Simulation") > 0 Then
                shp.TextFrame.TextRange.Text = "Flow Simulation: {MATERIAL_NAME}"
            End If
        End If
    Next shp
    
    ' Update material properties
    UpdateMaterialProperties
    
    ' Refresh charts
    RefreshCharts
    
    ' Update analysis
    UpdateAnalysisData
    
    MsgBox "FlowSim data updated successfully!", vbInformation
End Sub

Sub UpdateMaterialProperties()
    Dim slide As Slide
    Dim shp As Shape
    
    Set slide = ActivePresentation.Slides(2)
    
    ' Update density
    For Each shp In slide.Shapes
        If shp.HasTextFrame Then
            If InStr(shp.TextFrame.TextRange.Text, "Density:") > 0 Then
                shp.TextFrame.TextRange.Text = "Density: {DENSITY} g/cm³"
            ElseIf InStr(shp.TextFrame.TextRange.Text, "Friction:") > 0 Then
                shp.TextFrame.TextRange.Text = "Friction: {FRICTION}"
            ElseIf InStr(shp.TextFrame.TextRange.Text, "Elasticity:") > 0 Then
                shp.TextFrame.TextRange.Text = "Elasticity: {ELASTICITY}"
            End If
        End If
    Next shp
End Sub

Sub RefreshCharts()
    ' Refresh all charts in presentation
    Dim slide As Slide
    Dim shp As Shape
    
    For Each slide In ActivePresentation.Slides
        For Each shp In slide.Shapes
            If shp.HasChart Then
                shp.Chart.Refresh
            End If
        Next shp
    Next slide
End Sub

Sub ExportToExcel()
    ' Export data to Excel for further analysis
    Dim excelApp As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = True
    
    Set wb = excelApp.Workbooks.Add
    Set ws = wb.Worksheets(1)
    
    ' Add headers
    ws.Cells(1, 1).Value = "Property"
    ws.Cells(1, 2).Value = "Value"
    ws.Cells(1, 3).Value = "Unit"
    
    ' Add data
    ws.Cells(2, 1).Value = "Material Name"
    ws.Cells(2, 2).Value = "{MATERIAL_NAME}"
    
    ws.Cells(3, 1).Value = "Density"
    ws.Cells(3, 2).Value = {DENSITY}
    ws.Cells(3, 3).Value = "g/cm³"
    
    ws.Cells(4, 1).Value = "Friction"
    ws.Cells(4, 2).Value = {FRICTION}
    
    ws.Cells(5, 1).Value = "Elasticity"
    ws.Cells(5, 2).Value = {ELASTICITY}
    
    ws.Cells(6, 1).Value = "Flow Rate"
    ws.Cells(6, 2).Value = {FLOW_RATE}
    ws.Cells(6, 3).Value = "u/s"
    
    ' Format the table
    ws.Range("A1:C6").Borders.LineStyle = 1
    ws.Range("A1:C1").Font.Bold = True
    ws.Columns("A:C").AutoFit
    
    MsgBox "Data exported to Excel successfully!", vbInformation
    Exit Sub
    
ErrorHandler:
    MsgBox "Error exporting to Excel: " & Err.Description, vbExclamation
End Sub

' Run this macro when presentation opens
Sub Auto_Open()
    UpdateFlowSimData
End Sub`;
    }
    
    getAdvancedVBATemplate() {
        return `' FlowSim Material Studio - Advanced Engineering Macro
' Generated on {EXPORT_DATE}
' Material: {MATERIAL_NAME}
' User Mode: {USER_MODE}

Option Explicit

' Constants
Const SIMULATION_INTERVAL As Integer = 5 ' seconds
Const DATA_SHEET_NAME As String = "FlowSim_Data"
Const CHART_UPDATE As Boolean = True

' Global variables
Dim simulationTimer As Date
Dim dataPoints As Collection

Sub InitializeFlowSim()
    ' Initialize simulation data
    Set dataPoints = New Collection
    
    ' Create data worksheet if it doesn't exist
    CreateDataSheet
    
    ' Set up automatic updates
    simulationTimer = Now
    
    ' Update all slides
    UpdateAllSlides
    
    MsgBox "FlowSim Engineering Mode initialized", vbInformation
End Sub

Sub UpdateAllSlides()
    Dim slide As Slide
    Dim slideIndex As Integer
    
    For Each slide In ActivePresentation.Slides
        Select Case slideIndex
            Case 1 ' Title
                UpdateTitleSlide slide
            Case 2 ' Properties
                UpdatePropertiesSlide slide
            Case 3 ' Simulation
                UpdateSimulationSlide slide
            Case 4 ' Analysis
                UpdateAnalysisSlide slide
            Case 5 ' Recommendations
                UpdateRecommendationsSlide slide
            Case 6 ' Data
                UpdateDataSlide slide
            Case 7 ' Export
                UpdateExportSlide slide
        End Select
        
        slideIndex = slideIndex + 1
    Next slide
End Sub

Sub UpdateTitleSlide(slide As Slide)
    Dim shp As Shape
    
    For Each shp In slide.Shapes
        If shp.HasTextFrame Then
            Dim text As String
            text = shp.TextFrame.TextRange.Text
            
            If InStr(text, "Flow Simulation") > 0 Then
                shp.TextFrame.TextRange.Text = "ENGINEERING ANALYSIS: " & "{MATERIAL_NAME}"
                shp.TextFrame.TextRange.Font.Color = RGB(52, 152, 219)
            End If
        End If
    Next shp
End Sub

Sub UpdatePropertiesSlide(slide As Slide)
    ' Update with engineering data
    Dim properties As Variant
    properties = Array(
        Array("Density", {DENSITY}, "g/cm³", "High" & IIf({DENSITY} > 5, " (Heavy)", "")),
        Array("Friction", {FRICTION}, "μ", IIf({FRICTION} > 0.6, "High", IIf({FRICTION} < 0.3, "Low", "Medium"))),
        Array("Elasticity", {ELASTICITY}, "e", IIf({ELASTICITY} > 0.7, "Elastic", IIf({ELASTICITY} < 0.3, "Brittle", "Ductile"))),
        Array("Flow Rate", {FLOW_RATE}, "u/s", IIf({FLOW_RATE} > 7, "Excellent", IIf({FLOW_RATE} > 4, "Adequate", "Poor")))
    )
    
    ' Find and update property shapes
    Dim shp As Shape
    Dim i As Integer
    
    For Each shp In slide.Shapes
        If shp.HasTextFrame Then
            For i = 0 To UBound(properties)
                If InStr(shp.TextFrame.TextRange.Text, properties(i)(0)) > 0 Then
                    shp.TextFrame.TextRange.Text = properties(i)(0) & ": " & _
                        Format(properties(i)(1), "0.00") & " " & properties(i)(2) & _
                        " (" & properties(i)(3) & ")"
                End If
            Next i
        End If
    Next shp
End Sub

Sub RunEngineeringAnalysis()
    ' Perform advanced engineering calculations
    Dim density As Double
    Dim friction As Double
    Dim elasticity As Double
    
    density = {DENSITY}
    friction = {FRICTION}
    elasticity = {ELASTICITY}
    
    ' Calculate derived properties
    Dim flowIndex As Double
    Dim stabilityFactor As Double
    Dim efficiency As Double
    
    flowIndex = density * friction * 100
    stabilityFactor = (elasticity / friction) * 50
    efficiency = (flowIndex / stabilityFactor) * 100
    
    ' Output results
    Debug.Print "=== ENGINEERING ANALYSIS ==="
    Debug.Print "Material: " & "{MATERIAL_NAME}"
    Debug.Print "Flow Index: " & Format(flowIndex, "0.00")
    Debug.Print "Stability Factor: " & Format(stabilityFactor, "0.00")
    Debug.Print "System Efficiency: " & Format(efficiency, "0.00") & "%"
    Debug.Print "=========================="
    
    ' Update slide with results
    UpdateEngineeringResults flowIndex, stabilityFactor, efficiency
End Sub

Sub CreateDataSheet()
    ' Create Excel data sheet for logging
    On Error Resume Next
    
    Dim xlApp As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    
    ' Try to open existing workbook
    Set wb = xlApp.Workbooks.Open(Environ("USERPROFILE") & "\Documents\FlowSim_Data.xlsx")
    
    If wb Is Nothing Then
        ' Create new workbook
        Set wb = xlApp.Workbooks.Add
    End If
    
    ' Find or create worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(DATA_SHEET_NAME)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add
        ws.Name = DATA_SHEET_NAME
        
        ' Create headers
        ws.Cells(1, 1).Value = "Timestamp"
        ws.Cells(1, 2).Value = "Material"
        ws.Cells(1, 3).Value = "Density"
        ws.Cells(1, 4).Value = "Friction"
        ws.Cells(1, 5).Value = "Elasticity"
        ws.Cells(1, 6).Value = "Flow Rate"
        ws.Cells(1, 7).Value = "Flow Index"
        ws.Cells(1, 8).Value = "Stability"
    End If
    
    ' Add current data
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ws.Cells(lastRow, 1).Value = Now
    ws.Cells(lastRow, 2).Value = "{MATERIAL_NAME}"
    ws.Cells(lastRow, 3).Value = {DENSITY}
    ws.Cells(lastRow, 4).Value = {FRICTION}
    ws.Cells(lastRow, 5).Value = {ELASTICITY}
    ws.Cells(lastRow, 6).Value = {FLOW_RATE}
    
    ' Save and close
    wb.SaveAs Environ("USERPROFILE") & "\Documents\FlowSim_Data.xlsx"
    wb.Close
    xlApp.Quit
    
    Set ws = Nothing
    Set wb = Nothing
    Set xlApp = Nothing
End Sub

Sub GenerateEngineeringReport()
    ' Generate comprehensive engineering report
    Dim reportText As String
    
    reportText = "=== FLOWSIM ENGINEERING REPORT ===" & vbCrLf & vbCrLf
    reportText = reportText & "Material: " & "{MATERIAL_NAME}" & vbCrLf
    reportText = reportText & "Analysis Date: " & Format(Now, "yyyy-mm-dd HH:MM:ss") & vbCrLf & vbCrLf
    
    reportText = reportText & "PROPERTIES:" & vbCrLf
    reportText = reportText & "  Density: " & {DENSITY} & " g/cm³" & vbCrLf
    reportText = reportText & "  Friction: " & {FRICTION} & vbCrLf
    reportText = reportText & "  Elasticity: " & {ELASTICITY} & vbCrLf & vbCrLf
    
    reportText = reportText & "FLOW CHARACTERISTICS:" & vbCrLf
    reportText = reportText & "  Flow Rate: " & {FLOW_RATE} & " u/s" & vbCrLf
    reportText = reportText & "  Classification: " & IIf({FLOW_RATE} > 7, "Excellent", _
                                                     IIf({FLOW_RATE} > 4, "Adequate", "Poor")) & vbCrLf & vbCrLf
    
    reportText = reportText & "RECOMMENDATIONS:" & vbCrLf
    
    If {DENSITY} > 5 Then
        reportText = reportText & "  - Use heavy-duty equipment" & vbCrLf
    End If
    
    If {FRICTION} > 0.6 Then
        reportText = reportText & "  - Consider surface treatment" & vbCrLf
    End If
    
    If {FLOW_RATE} < 4 Then
        reportText = reportText & "  - Review flow path design" & vbCrLf
    End If
    
    ' Display report
    MsgBox reportText, vbInformation, "Engineering Report"
    
    ' Also save to file
    SaveReportToFile reportText
End Sub

Sub SaveReportToFile(reportText As String)
    Dim fso As Object
    Dim file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim filePath As String
    filePath = Environ("USERPROFILE") & "\Documents\FlowSim_Report_" & _
               Format(Now, "yyyymmdd_HHMMss") & ".txt"
    
    Set file = fso.CreateTextFile(filePath, True)
    file.Write reportText
    file.Close
    
    MsgBox "Report saved to: " & filePath, vbInformation
End Sub

' Automatic update on slide show start
Sub OnSlideShowPageChange()
    If Now > DateAdd("s", SIMULATION_INTERVAL, simulationTimer) Then
        UpdateAllSlides
        simulationTimer = Now
    End If
End Sub`;
    }
    
    getSimulationVBATemplate() {
        return `' FlowSim Material Studio - Educational Simulation Macro
' Generated on {EXPORT_DATE}
' Material: {MATERIAL_NAME}
' Educational Mode

Option Explicit

' Educational constants
Const MAX_STUDENTS As Integer = 40
Const QUIZ_QUESTIONS As Integer = 5

' Student data structure
Type StudentRecord
    Name As String
    Score As Integer
    Completed As Boolean
    Feedback As String
End Type

Dim students(1 To MAX_STUDENTS) As StudentRecord
Dim studentCount As Integer

Sub InitializeEducationalMode()
    ' Setup educational presentation
    studentCount = 0
    
    ' Add interactive elements
    AddInteractiveQuestions
    AddSimulationControls
    AddQuizSection
    
    ' Update educational content
    UpdateEducationalContent
    
    MsgBox "Educational mode initialized. Ready for classroom use.", vbInformation
End Sub

Sub UpdateEducationalContent()
    Dim slide As Slide
    
    For Each slide In ActivePresentation.Slides
        ' Add educational notes
        AddEducationalNotes slide
        
        ' Update learning objectives
        UpdateLearningObjectives slide
        
        ' Add discussion questions
        AddDiscussionQuestions slide
    Next slide
End Sub

Sub AddInteractiveQuestions()
    ' Add interactive quiz questions about the material
    Dim questions As Variant
    questions = Array(
        "What is the primary factor affecting flow rate?",
        "How does density influence material flow?",
        "What practical applications use this material?",
        "How would you improve the flow of this material?",
        "What safety considerations are important?"
    )
    
    Dim answers As Variant
    answers = Array(
        "Friction coefficient",
        "Higher density requires more energy to move",
        "Manufacturing, construction, agriculture",
        "Reduce friction, increase slope, add vibration",
        "Dust control, static electricity, containment"
    )
    
    ' Create quiz slides
    Dim i As Integer
    For i = 0 To UBound(questions)
        CreateQuizSlide i + 1, questions(i), answers(i)
    Next i
End Sub

Sub CreateQuizSlide(questionNum As Integer, question As String, answer As String)
    Dim newSlide As Slide
    Set newSlide = ActivePresentation.Slides.Add(ActivePresentation.Slides.Count + 1, ppLayoutTitleOnly)
    
    newSlide.Shapes.Title.TextFrame.TextRange.Text = "Question " & questionNum
    
    ' Add question
    Dim questionBox As Shape
    Set questionBox = newSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 100, 600, 100)
    questionBox.TextFrame.TextRange.Text = question
    questionBox.TextFrame.TextRange.Font.Size = 18
    
    ' Add answer (initially hidden)
    Dim answerBox As Shape
    Set answerBox = newSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 250, 600, 100)
    answerBox.TextFrame.TextRange.Text = "Answer: " & answer
    answerBox.TextFrame.TextRange.Font.Size = 16
    answerBox.TextFrame.TextRange.Font.Color = RGB(52, 152, 219)
    answerBox.Visible = msoFalse
    
    ' Add reveal button
    Dim revealBtn As Shape
    Set revealBtn = newSlide.Shapes.AddShape(msoShapeRoundedRectangle, 250, 400, 200, 50)
    revealBtn.TextFrame.TextRange.Text = "Reveal Answer"
    revealBtn.TextFrame.TextRange.Font.Size = 14
    revealBtn.Fill.ForeColor.RGB = RGB(46, 204, 113)
    
    ' Add animation to reveal answer
    Dim effect As Effect
    Set effect = newSlide.TimeLine.MainSequence.AddEffect(answerBox, msoAnimEffectFade)
    effect.Timing.TriggerType = msoAnimTriggerOnShapeClick
    effect.Timing.TriggerShape = revealBtn
End Sub

Sub AddStudentRecord(studentName As String, score As Integer)
    If studentCount < MAX_STUDENTS Then
        studentCount = studentCount + 1
        
        students(studentCount).Name = studentName
        students(studentCount).Score = score
        students(studentCount).Completed = True
        students(studentCount).Feedback = GenerateFeedback(score)
        
        ' Update leaderboard
        UpdateLeaderboard
    End If
End Sub

Function GenerateFeedback(score As Integer) As String
    Select Case score
        Case Is >= 90
            GenerateFeedback = "Excellent understanding of material flow concepts"
        Case Is >= 70
            GenerateFeedback = "Good grasp of basic principles, review advanced topics"
        Case Is >= 50
            GenerateFeedback = "Basic understanding achieved, practice recommended"
        Case Else
            GenerateFeedback = "Review fundamental concepts and retry quiz"
    End Select
End Function

Sub UpdateLeaderboard()
    ' Sort students by score
    Dim i As Integer, j As Integer
    Dim temp As StudentRecord
    
    For i = 1 To studentCount - 1
        For j = i + 1 To studentCount
            If students(i).Score < students(j).Score Then
                temp = students(i)
                students(i) = students(j)
                students(j) = temp
            End If
        Next j
    Next i
    
    ' Update leaderboard slide
    UpdateLeaderboardSlide
End Sub

Sub UpdateLeaderboardSlide()
    On Error Resume Next
    
    Dim leaderboardSlide As Slide
    Set leaderboardSlide = ActivePresentation.Slides("Leaderboard")
    
    If leaderboardSlide Is Nothing Then
        Set leaderboardSlide = ActivePresentation.Slides.Add(ActivePresentation.Slides.Count + 1, ppLayoutBlank)
        leaderboardSlide.Name = "Leaderboard"
        
        ' Add title
        Dim title As Shape
        Set title = leaderboardSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 50)
        title.TextFrame.TextRange.Text = "Class Leaderboard"
        title.TextFrame.TextRange.Font.Size = 32
        title.TextFrame.TextRange.Font.Bold = True
    End If
    
    ' Clear existing content (except title)
    Dim shp As Shape
    For Each shp In leaderboardSlide.Shapes
        If shp.Name <> "Title" Then
            shp.Delete
        End If
    Next shp
    
    ' Add leaderboard table
    Dim tableTop As Integer: tableTop = 120
    Dim rowHeight As Integer: rowHeight = 30
    
    ' Headers
    AddTableRow leaderboardSlide, 50, tableTop, Array("Rank", "Student", "Score", "Feedback"), True
    
    ' Student rows
    Dim i As Integer
    For i = 1 To studentCount
        If i > 10 Then Exit For ' Show top 10 only
        
        AddTableRow leaderboardSlide, 50, tableTop + (i * rowHeight), _
            Array(i, students(i).Name, students(i).Score, students(i).Feedback), False
    Next i
End Sub

Sub AddTableRow(slide As Slide, left As Integer, top As Integer, data As Variant, isHeader As Boolean)
    Dim cellWidth As Integer: cellWidth = 150
    Dim i As Integer
    
    For i = 0 To UBound(data)
        Dim cell As Shape
        Set cell = slide.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                                           left + (i * cellWidth), top, cellWidth, 30)
        
        cell.TextFrame.TextRange.Text = data(i)
        
        If isHeader Then
            cell.TextFrame.TextRange.Font.Bold = True
            cell.Fill.ForeColor.RGB = RGB(52, 152, 219)
            cell.TextFrame.TextRange.Font.Color = RGB(255, 255, 255)
        End If
        
        cell.Line.ForeColor.RGB = RGB(200, 200, 200)
    Next i
End Sub

Sub GenerateClassReport()
    Dim report As String
    Dim avgScore As Double
    Dim totalScore As Integer
    Dim i As Integer
    
    ' Calculate statistics
    For i = 1 To studentCount
        totalScore = totalScore + students(i).Score
    Next i
    
    If studentCount > 0 Then
        avgScore = totalScore / studentCount
    End If
    
    ' Generate report
    report = "=== CLASS PERFORMANCE REPORT ===" & vbCrLf & vbCrLf
    report = report & "Material Studied: " & "{MATERIAL_NAME}" & vbCrLf
    report = report & "Date: " & Format(Now, "yyyy-mm-dd") & vbCrLf & vbCrLf
    
    report = report & "CLASS STATISTICS:" & vbCrLf
    report = report & "  Students: " & studentCount & vbCrLf
    report = report & "  Average Score: " & Format(avgScore, "0.0") & vbCrLf
    report = report & "  Highest Score: " & IIf(studentCount > 0, students(1).Score, "N/A") & vbCrLf
    report = report & "  Completion Rate: 100%" & vbCrLf & vbCrLf
    
    report = report & "MATERIAL PROPERTIES (for reference):" & vbCrLf
    report = report & "  Density: " & {DENSITY} & " g/cm³" & vbCrLf
    report = report & "  Friction: " & {FRICTION} & vbCrLf
    report = report & "  Elasticity: " & {ELASTICITY} & vbCrLf
    report = report & "  Flow Rate: " & {FLOW_RATE} & " u/s" & vbCrLf
    
    ' Display report
    MsgBox report, vbInformation, "Class Report"
    
    ' Save report
    SaveClassReport report
End Sub

Sub SaveClassReport(reportText As String)
    Dim fso As Object
    Dim file As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim filePath As String
    filePath = Environ("USERPROFILE") & "\Documents\FlowSim_ClassReport_" & _
               Format(Now, "yyyymmdd") & ".txt"
    
    Set file = fso.CreateTextFile(filePath, True)
    file.Write reportText
    file.Close
    
    ' Also copy to slide
    AddReportToPresentation reportText
End Sub

Sub AddReportToPresentation(reportText As String)
    Dim reportSlide As Slide
    Set reportSlide = ActivePresentation.Slides.Add(ActivePresentation.Slides.Count + 1, ppLayoutTitleOnly)
    
    reportSlide.Shapes.Title.TextFrame.TextRange.Text = "Class Report Summary"
    
    Dim reportBox As Shape
    Set reportBox = reportSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 100, 600, 400)
    
    ' Split report into lines and add
    Dim lines As Variant
    lines = Split(reportText, vbCrLf)
    
    Dim i As Integer
    For i = 0 To UBound(lines)
        reportBox.TextFrame.TextRange.InsertAfter lines(i) & vbCrLf
    Next i
    
    reportBox.TextFrame.TextRange.Font.Size = 12
    reportBox.TextFrame.TextRange.Font.Name = "Consolas"
End Sub

' Animation for educational demonstration
Sub AnimateFlowSimulation()
    Dim slide As Slide
    Dim shp As Shape
    
    Set slide = ActivePresentation.Slides(3) ' Simulation slide
    
    ' Find and animate flow arrows
    For Each shp In slide.Shapes
        If shp.Name Like "Arrow_*" Then
            Dim effect As Effect
            Set effect = slide.TimeLine.MainSequence.AddEffect(shp, msoAnimEffectPathRight)
            
            effect.Timing.Duration = 3
            effect.Timing.RepeatCount = 3
        End If
    Next shp
End Sub`;
    }
    
    getReportVBATemplate() {
        return `' FlowSim Material Studio - Presentation Report Macro
' Generated on {EXPORT_DATE}
' Material: {MATERIAL_NAME}
' Presentation Mode

Option Explicit

' Presentation settings
Const AUTO_ADVANCE As Boolean = True
Const ADVANCE_INTERVAL As Integer = 30 ' seconds
Const SHOW_NOTES As Boolean = False

' Timer for auto-advance
Dim presentationTimer As Date
Dim currentSlide As Integer

Sub InitializePresentationMode()
    ' Setup for professional presentation
    presentationTimer = Now
    currentSlide = 1
    
    ' Apply presentation theme
    ApplyPresentationTheme
    
    ' Setup auto-advance if enabled
    If AUTO_ADVANCE Then
        SetupAutoAdvance
    End If
    
    ' Hide notes if not needed
    If Not SHOW_NOTES Then
        HidePresenterNotes
    End If
    
    ' Add navigation controls
    AddNavigationControls
    
    MsgBox "Presentation mode activated. Press F5 to start slideshow.", vbInformation
End Sub

Sub ApplyPresentationTheme()
    ' Apply professional styling
    Dim slide As Slide
    Dim shp As Shape
    
    For Each slide In ActivePresentation.Slides
        ' Update font sizes for presentation
        For Each shp In slide.Shapes
            If shp.HasTextFrame Then
                Select Case shp.Type
                    Case msoPlaceholder
                        ' Title placeholder
                        If shp.PlaceholderFormat.Type = ppPlaceholderTitle Then
                            shp.TextFrame.TextRange.Font.Size = 44
                            shp.TextFrame.TextRange.Font.Bold = True
                            shp.TextFrame.TextRange.Font.Color = RGB(52, 152, 219)
                        
                        ' Content placeholder
                        ElseIf shp.PlaceholderFormat.Type = ppPlaceholderBody Then
                            shp.TextFrame.TextRange.Font.Size = 28
                        End If
                    
                    ' Regular text boxes
                    Case msoTextBox
                        If shp.TextFrame.TextRange.Font.Size < 24 Then
                            shp.TextFrame.TextRange.Font.Size = 24
                        End If
                End Select
            End If
        Next shp
        
        ' Add slide numbers
        AddSlideNumber slide
    Next slide
End Sub

Sub AddSlideNumber(slide As Slide)
    Dim slideNum As Shape
    
    ' Check if slide number already exists
    For Each slideNum In slide.Shapes
        If slideNum.Type = msoPlaceholder And _
           slideNum.PlaceholderFormat.Type = ppPlaceholderSlideNumber Then
            Exit Sub
        End If
    Next slideNum
    
    ' Add slide number
    Set slideNum = slide.Shapes.AddPlaceholder(ppPlaceholderSlideNumber, _
                                               -1, -1, -1, -1)
    
    ' Format slide number
    slideNum.TextFrame.TextRange.Text = slide.SlideIndex
    slideNum.TextFrame.TextRange.Font.Size = 14
    slideNum.TextFrame.TextRange.Font.Color = RGB(150, 150, 150)
    slideNum.Left = ActivePresentation.PageSetup.SlideWidth - 100
    slideNum.Top = ActivePresentation.PageSetup.SlideHeight - 50
End Sub

Sub SetupAutoAdvance()
    Dim slide As Slide
    
    For Each slide In ActivePresentation.Slides
        slide.SlideShowTransition.AdvanceOnTime = True
        slide.SlideShowTransition.AdvanceTime = ADVANCE_INTERVAL
    Next slide
End Sub

Sub AddNavigationControls()
    ' Add navigation buttons to key slides
    AddNavigationToSlide 1 ' Title slide
    AddNavigationToSlide 3 ' Simulation slide
    AddNavigationToSlide 5 ' Recommendations slide
End Sub

Sub AddNavigationToSlide(slideIndex As Integer)
    Dim slide As Slide
    Dim navButton As Shape
    
    Set slide = ActivePresentation.Slides(slideIndex)
    
    ' Add home button
    Set navButton = slide.Shapes.AddShape(msoShapeActionButtonCustom, 20, 20, 40, 40)
    navButton.TextFrame.TextRange.Text = "🏠"
    navButton.TextFrame.TextRange.Font.Size = 20
    navButton.ActionSettings(ppMouseClick).Action = ppActionRunMacro
    navButton.ActionSettings(ppMouseClick).Run = "GoToSlide:1"
    navButton.Name = "Nav_Home_" & slideIndex
    
    ' Add next button
    Set navButton = slide.Shapes.AddShape(msoShapeActionButtonCustom, 70, 20, 40, 40)
    navButton.TextFrame.TextRange.Text = "▶"
    navButton.TextFrame.TextRange.Font.Size = 20
    navButton.ActionSettings(ppMouseClick).Action = ppActionRunMacro
    navButton.ActionSettings(ppMouseClick).Run = "GoToSlide:" & (slideIndex + 1)
    navButton.Name = "Nav_Next_" & slideIndex
    
    ' Add previous button (if not first slide)
    If slideIndex > 1 Then
        Set navButton = slide.Shapes.AddShape(msoShapeActionButtonCustom, 120, 20, 40, 40)
        navButton.TextFrame.TextRange.Text = "◀"
        navButton.TextFrame.TextRange.Font.Size = 20
        navButton.ActionSettings(ppMouseClick).Action = ppActionRunMacro
        navButton.ActionSettings(ppMouseClick).Run = "GoToSlide:" & (slideIndex - 1)
        navButton.Name = "Nav_Prev_" & slideIndex
    End If
    
    ' Add menu button
    Set navButton = slide.Shapes.AddShape(msoShapeActionButtonCustom, _
                                          ActivePresentation.PageSetup.SlideWidth - 60, 20, 40, 40)
    navButton.TextFrame.TextRange.Text = "☰"
    navButton.TextFrame.TextRange.Font.Size = 20
    navButton.ActionSettings(ppMouseClick).Action = ppActionRunMacro
    navButton.ActionSettings(ppMouseClick).Run = "ShowSlideMenu"
    navButton.Name = "Nav_Menu_" & slideIndex
End Sub

Sub GoToSlide(slideIndex As Integer)
    SlideShowWindows(1).View.GotoSlide slideIndex
End Sub

Sub ShowSlideMenu()
    ' Create slide menu overlay
    Dim menuSlide As Slide
    Set menuSlide = ActivePresentation.Slides.Add(ActivePresentation.Slides.Count + 1, ppLayoutBlank)
    
    menuSlide.FollowMasterBackground = False
    menuSlide.Background.Fill.ForeColor.RGB = RGB(0, 0, 0)
    menuSlide.Background.Fill.Transparency = 0.3
    
    ' Add menu title
    Dim title As Shape
    Set title = menuSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 50, 600, 50)
    title.TextFrame.TextRange.Text = "Slide Menu"
    title.TextFrame.TextRange.Font.Size = 36
    title.TextFrame.TextRange.Font.Color = RGB(255, 255, 255)
    title.TextFrame.TextRange.Font.Bold = True
    
    ' Add slide thumbnails
    Dim i As Integer
    Dim thumbTop As Integer: thumbTop = 120
    Dim thumbLeft As Integer: thumbLeft = 50
    
    For i = 1 To ActivePresentation.Slides.Count - 1 ' Exclude menu slide
        If i <= 6 Then ' Show first 6 slides
            AddSlideThumbnail menuSlide, i, thumbLeft, thumbTop
            thumbLeft = thumbLeft + 200
            
            If thumbLeft > 450 Then
                thumbLeft = 50
                thumbTop = thumbTop + 150
            End If
        End If
    Next i
    
    ' Add close button
    Dim closeBtn As Shape
    Set closeBtn = menuSlide.Shapes.AddShape(msoShapeRoundedRectangle, _
                                             ActivePresentation.PageSetup.SlideWidth - 100, 20, 80, 40)
    closeBtn.TextFrame.TextRange.Text = "Close"
    closeBtn.TextFrame.TextRange.Font.Size = 14
    closeBtn.Fill.ForeColor.RGB = RGB(231, 76, 60)
    closeBtn.TextFrame.TextRange.Font.Color = RGB(255, 255, 255)
    closeBtn.ActionSettings(ppMouseClick).Hyperlink.Address = ""
    closeBtn.ActionSettings(ppMouseClick).Action = ppActionLastSlide
End Sub

Sub AddSlideThumbnail(menuSlide As Slide, slideIndex As Integer, left As Integer, top As Integer)
    Dim thumb As Shape
    Dim slide As Slide
    
    Set slide = ActivePresentation.Slides(slideIndex)
    
    ' Create thumbnail background
    Set thumb = menuSlide.Shapes.AddShape(msoShapeRoundedRectangle, left, top, 180, 120)
    thumb.Fill.ForeColor.RGB = RGB(255, 255, 255)
    thumb.Fill.Transparency = 0.1
    
    ' Add slide number
    Dim numText As Shape
    Set numText = menuSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, left + 10, top + 10, 30, 30)
    numText.TextFrame.TextRange.Text = slideIndex
    numText.TextFrame.TextRange.Font.Size = 16
    numText.TextFrame.TextRange.Font.Color = RGB(255, 255, 255)
    numText.TextFrame.TextRange.Font.Bold = True
    
    ' Add slide title (truncated)
    Dim titleText As Shape
    Set titleText = menuSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, left + 50, top + 10, 120, 50)
    
    Dim slideTitle As String
    If slide.Shapes.Title Is Nothing Then
        slideTitle = "Slide " & slideIndex
    Else
        slideTitle = slide.Shapes.Title.TextFrame.TextRange.Text
        If Len(slideTitle) > 20 Then
            slideTitle = Left(slideTitle, 20) & "..."
        End If
    End If
    
    titleText.TextFrame.TextRange.Text = slideTitle
    titleText.TextFrame.TextRange.Font.Size = 12
    titleText.TextFrame.TextRange.Font.Color = RGB(255, 255, 255)
    
    ' Make thumbnail clickable
    thumb.ActionSettings(ppMouseClick).Hyperlink.SubAddress = slideIndex & "," & slide.Name
End Sub

Sub GenerateExecutiveSummary()
    ' Create executive summary slide
    Dim summarySlide As Slide
    Set summarySlide = ActivePresentation.Slides.Add(2, ppLayoutTitleOnly) ' Insert after title
    
    summarySlide.Shapes.Title.TextFrame.TextRange.Text = "Executive Summary"
    
    ' Add summary content
    Dim summaryBox As Shape
    Set summaryBox = summarySlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 100, 600, 300)
    
    Dim summaryText As String
    summaryText = "MATERIAL: " & "{MATERIAL_NAME}" & vbCrLf & vbCrLf
    
    summaryText = summaryText & "KEY FINDINGS:" & vbCrLf
    summaryText = summaryText & "• Flow rate: " & {FLOW_RATE} & " u/s" & vbCrLf
    summaryText = summaryText & "• Density classification: " & _
                  IIf({DENSITY} > 5, "High", IIf({DENSITY} > 2, "Medium", "Low")) & vbCrLf
    summaryText = summaryText & "• Flow characteristics: " & _
                  IIf({FLOW_RATE} > 7, "Excellent", IIf({FLOW_RATE} > 4, "Adequate", "Needs improvement")) & vbCrLf & vbCrLf
    
    summaryText = summaryText & "RECOMMENDATIONS:" & vbCrLf
    
    If {DENSITY} > 5 Then
        summaryText = summaryText & "• Use heavy-duty conveying equipment" & vbCrLf
    End If
    
    If {FRICTION} > 0.6 Then
        summaryText = summaryText & "• Consider surface treatment or lubrication" & vbCrLf
    End If
    
    If {FLOW_RATE} < 4 Then
        summaryText = summaryText & "• Review system design for better flow" & vbCrLf
    End If
    
    summaryBox.TextFrame.TextRange.Text = summaryText
    summaryBox.TextFrame.TextRange.Font.Size = 20
    summaryBox.TextFrame.TextRange.Font.Color = RGB(44, 62, 80)
    
    ' Add key metrics visualization
    AddMetricsVisualization summarySlide
End Sub

Sub AddMetricsVisualization(slide As Slide)
    ' Add a simple metrics visualization
    Dim metricsLeft As Integer: metricsLeft = 50
    Dim metricsTop As Integer: metricsTop = 400
    
    ' Density gauge
    AddMetricGauge slide, metricsLeft, metricsTop, "Density", {DENSITY}, 10, "g/cm³"
    
    ' Flow rate gauge
    AddMetricGauge slide, metricsLeft + 200, metricsTop, "Flow Rate", {FLOW_RATE}, 10, "u/s"
    
    ' Friction gauge
    AddMetricGauge slide, metricsLeft + 400, metricsTop, "Friction", {FRICTION}, 1, "μ"
End Sub

Sub AddMetricGauge(slide As Slide, left As Integer, top As Integer, _
                   label As String, value As Double, maxValue As Double, unit As String)
    ' Gauge background
    Dim gaugeBack As Shape
    Set gaugeBack = slide.Shapes.AddShape(msoShapeChord, left, top, 150, 150)
    gaugeBack.Fill.ForeColor.RGB = RGB(230, 230, 230)
    gaugeBack.Line.ForeColor.RGB = RGB(200, 200, 200)
    
    ' Gauge fill (based on value)
    Dim fillPercent As Double
    fillPercent = value / maxValue
    If fillPercent > 1 Then fillPercent = 1
    
    Dim gaugeFill As Shape
    Set gaugeFill = slide.Shapes.AddShape(msoShapeChord, left, top, 150, 150)
    
    ' Color based on value
    Dim fillColor As Long
    If fillPercent > 0.7 Then
        fillColor = RGB(46, 204, 113) ' Green
    ElseIf fillPercent > 0.4 Then
        fillColor = RGB(241, 196, 15) ' Yellow
    Else
        fillColor = RGB(231, 76, 60) ' Red
    End If
    
    gaugeFill.Fill.ForeColor.RGB = fillColor
    gaugeFill.Adjustments(1) = 0.5 - (fillPercent * 0.5)
    gaugeFill.Adjustments(2) = 0.5 + (fillPercent * 0.5)
    
    ' Label
    Dim gaugeLabel As Shape
    Set gaugeLabel = slide.Shapes.AddTextbox(msoTextOrientationHorizontal, left, top + 160, 150, 30)
    gaugeLabel.TextFrame.TextRange.Text = label & ": " & Format(value, "0.00") & " " & unit
    gaugeLabel.TextFrame.TextRange.Font.Size = 12
    gaugeLabel.TextFrame.TextRange.Font.Color = RGB(100, 100, 100)
    gaugeLabel.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
End Sub

Sub ExportPresentationPackage()
    ' Create a complete presentation package
    Dim packagePath As String
    packagePath = Environ("USERPROFILE") & "\Documents\FlowSim_Presentation_" & _
                  Format(Now, "yyyymmdd_HHMMss")
    
    ' Create folder
    MkDir packagePath
    
    ' Save presentation
    ActivePresentation.SaveCopyAs packagePath & "\Presentation.pptx"
    
    ' Export slides as images
    ExportSlidesAsImages packagePath
    
    ' Create speaker notes
    CreateSpeakerNotes packagePath
    
    ' Create handouts
    CreateHandouts packagePath
    
    ' Create readme
    CreateReadmeFile packagePath
    
    MsgBox "Presentation package created at:" & vbCrLf & packagePath, vbInformation
End Sub

Sub ExportSlidesAsImages(folderPath As String)
    Dim slide As Slide
    Dim i As Integer
    
    For i = 1 To ActivePresentation.Slides.Count
        Set slide = ActivePresentation.Slides(i)
        slide.Export folderPath & "\Slide_" & Format(i, "00") & ".png", "PNG"
    Next i
End Sub

Sub CreateSpeakerNotes(folderPath As String)
    Dim notesText As String
    Dim slide As Slide
    Dim i As Integer
    
    notesText = "FLOWSIM PRESENTATION - SPEAKER NOTES" & vbCrLf & vbCrLf
    notesText = notesText & "Material: " & "{MATERIAL_NAME}" & vbCrLf
    notesText = notesText & "Generated: " & Format(Now, "yyyy-mm-dd HH:MM:ss") & vbCrLf & vbCrLf
    
    For i = 1 To ActivePresentation.Slides.Count
        Set slide = ActivePresentation.Slides(i)
        
        notesText = notesText & "=== SLIDE " & i & " ===" & vbCrLf
        
        ' Slide title
        If Not slide.Shapes.Title Is Nothing Then
            notesText = notesText & "Title: " & slide.Shapes.Title.TextFrame.TextRange.Text & vbCrLf
        End If
        
        ' Speaker notes
        If slide.HasNotesPage Then
            Dim notesPage As NotesPage
            Set notesPage = slide.NotesPage
            
            Dim notesShape As Shape
            For Each notesShape In notesPage.Shapes
                If notesShape.PlaceholderFormat.Type = ppPlaceholderBody Then
                    Dim noteText As String
                    noteText = notesShape.TextFrame.TextRange.Text
                    If Len(noteText) > 0 Then
                        notesText = notesText & "Notes: " & noteText & vbCrLf
                    End If
                End If
            Next notesShape
        End If
        
        ' Key points
        notesText = notesText & "Key Points:" & vbCrLf
        
        Select Case i
            Case 1
                notesText = notesText & "- Welcome audience" & vbCrLf
                notesText = notesText & "- Introduce FlowSim Material Studio" & vbCrLf
                notesText = notesText & "- State presentation objectives" & vbCrLf
            Case 2
                notesText = notesText & "- Present executive summary" & vbCrLf
                notesText = notesText & "- Highlight key findings" & vbCrLf
                notesText = notesText & "- Set expectations" & vbCrLf
            Case 3
                notesText = notesText & "- Explain material properties" & vbCrLf
                notesText = notesText & "- Discuss practical implications" & vbCrLf
                notesText = notesText & "- Relate to audience experience" & vbCrLf
        End Select
        
        notesText = notesText & vbCrLf
    Next i
    
    ' Save notes
    Dim fso As Object, file As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.CreateTextFile(folderPath & "\Speaker_Notes.txt", True)
    file.Write notesText
    file.Close
End Sub

' Timer event for presentation
Sub OnSlideShowNextBuild()
    If AUTO_ADVANCE Then
        If DateDiff("s", presentationTimer, Now) >= ADVANCE_INTERVAL Then
            SlideShowWindows(1).View.Next
            presentationTimer = Now
        End If
    End If
End Sub`;
    }
    
    // ====== UTILITY FUNCTIONS ======
    
    getRecommendationIcon(recommendation) {
        if (recommendation.toLowerCase().includes('density') || recommendation.toLowerCase().includes('heavy')) {
            return 'fas fa-weight-hanging';
        } else if (recommendation.toLowerCase().includes('friction') || recommendation.toLowerCase().includes('surface')) {
            return 'fas fa-magnet';
        } else if (recommendation.toLowerCase().includes('flow') || recommendation.toLowerCase().includes('design')) {
            return 'fas fa-tachometer-alt';
        } else if (recommendation.toLowerCase().includes('vibration') || recommendation.toLowerCase().includes('agitation')) {
            return 'fas fa-vibrate';
        } else {
            return 'fas fa-lightbulb';
        }
    }
    
    getRecommendationPriority(recommendation) {
        if (recommendation.toLowerCase().includes('critical') || recommendation.toLowerCase().includes('essential')) {
            return 'high';
        } else if (recommendation.toLowerCase().includes('consider') || recommendation.toLowerCase().includes('suggest')) {
            return 'medium';
        } else {
            return 'low';
        }
    }
    
    calculateMaterialFlowRate(material) {
        // Simplified flow rate calculation based on material properties
        if (!material.properties) return 0;
        
        const density = material.properties.density || 1;
        const friction = material.properties.friction || 0.5;
        const elasticity = material.properties.elasticity || 0.5;
        
        // Flow rate formula (simplified)
        return (10 / density) * (1 - friction) * elasticity;
    }
    
    getReportType(userMode) {
        switch (userMode) {
            case 'engineer': return 'Engineering Analysis Report';
            case 'educator': return 'Educational Presentation';
            case 'presenter': return 'Professional Presentation';
            default: return 'Analysis Report';
        }
    }
    
    // ====== FINALIZATION AND DOWNLOAD ======
    
    async finalizeExport(presentation, options) {
        if (this.officeJSLoaded && this.currentPresentation) {
            // Save using Office.js
            return await this.saveWithOfficeJS(presentation, options);
        } else {
            // Create download for simulated presentation
            return this.createDownload(presentation, options);
        }
    }
    
    async saveWithOfficeJS(presentation, options) {
        return new Promise((resolve) => {
            // In a real implementation, this would use Office.js API
            // to save the presentation
            
            setTimeout(() => {
                resolve({
                    success: true,
                    filename: presentation.name,
                    path: 'presentations/' + presentation.name,
                    slideCount: presentation.slideCount,
                    vbaIncluded: options.includeVBA
                });
            }, 2000);
        });
    }
    
    createDownload(presentation, options) {
        // Create a downloadable PowerPoint file
        // Note: This is a simplified simulation
        
        const exportData = {
            presentation: presentation,
            vbaCode: options.includeVBA ? this.vbaMacros[presentation.id] : null,
            metadata: {
                exportedAt: new Date().toISOString(),
                exportOptions: options,
                version: '1.0'
            }
        };
        
        // Create JSON file for demo (in production, would create actual .pptx)
        const jsonStr = JSON.stringify(exportData, null, 2);
        const blob = new Blob([jsonStr], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        
        const a = document.createElement('a');
        a.href = url;
        a.download = presentation.name.replace('.pptx', '.json');
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        
        return {
            success: true,
            filename: presentation.name,
            type: 'simulated',
            message: 'Simulated export completed. In production, this would create a .pptx file.'
        };
    }
    
    showExportSuccess(result, options) {
        // Update export stats
        this.updateExportStats();
        
        // Show success message
        const message = options.autoOpen ? 
            `✅ Presentation exported successfully! (${result.slideCount} slides)` :
            `✅ Presentation saved: ${result.filename}`;
        
        showNotification(message, 'success');
        
        // If auto-open is enabled and we have Office.js
        if (options.autoOpen && this.officeJSLoaded) {
            setTimeout(() => {
                showNotification('Opening presentation in PowerPoint...', 'info');
            }, 1000);
        }
    }
    
    updateExportStats() {
        // Update export statistics in UI
        const statsElement = document.querySelector('.stat-exports-today');
        if (statsElement) {
            const currentCount = parseInt(statsElement.textContent) || 0;
            statsElement.textContent = currentCount + 1;
        }
        
        // Update localStorage
        const today = new Date().toDateString();
        const storageKey = `flowsim_exports_${today}`;
        const todayExports = parseInt(localStorage.getItem(storageKey) || '0');
        localStorage.setItem(storageKey, (todayExports + 1).toString());
    }
    
    // ====== SETTINGS MANAGEMENT ======
    
    loadSettings() {
        const savedSettings = localStorage.getItem('flowsim_export_settings');
        if (savedSettings) {
            try {
                this.exportSettings = { ...this.exportSettings, ...JSON.parse(savedSettings) };
                console.log('📋 Loaded export settings');
            } catch (error) {
                console.error('Failed to load settings:', error);
            }
        }
    }
    
    saveSettings() {
        try {
            localStorage.setItem('flowsim_export_settings', JSON.stringify(this.exportSettings));
            console.log('💾 Export settings saved');
        } catch (error) {
            console.error('Failed to save settings:', error);
        }
    }
    
    updateSetting(key, value) {
        this.exportSettings[key] = value;
        this.saveSettings();
    }
    
    // ====== EVENT LISTENERS ======
    
    setupEventListeners() {
        // Listen for export modal events
        document.addEventListener('click', (e) => {
            if (e.target.closest('.btn-export-ppt')) {
                this.openExportModal();
            }
            
            if (e.target.closest('.btn-ppt-integration')) {
                this.exportToPowerPoint();
            }
        });
        
        // Listen for setting changes
        document.addEventListener('change', (e) => {
            if (e.target.closest('.export-settings select')) {
                const select = e.target;
                const key = select.id.replace('export', '').toLowerCase();
                this.updateSetting(key, select.value);
            }
        });
        
        // Listen for VBA copy/download
        document.addEventListener('click', (e) => {
            if (e.target.closest('[data-action="copy-vba"]')) {
                this.copyVBACode();
            }
            
            if (e.target.closest('[data-action="download-vba"]')) {
                this.downloadVBAMacro();
            }
        });
    }
    
    // ====== MODAL UI ======
    
    openExportModal() {
        const modal = this.createExportModal();
        document.body.appendChild(modal);
    }
    
    createExportModal() {
        const modal = document.createElement('div');
        modal.className = 'modal visible';
        modal.innerHTML = `
            <div class="modal-content export-modal">
                <div class="modal-header">
                    <h3><i class="fas fa-file-export"></i> PowerPoint Export</h3>
                    <button class="btn-close" onclick="this.closest('.modal').remove()">
                        <i class="fas fa-times"></i>
                    </button>
                </div>
                
                <div class="modal-body">
                    <div class="connection-status">
                        <div class="status-indicator ${this.isConnected ? 'connected' : ''}"></div>
                        <span>${this.isConnected ? 'Connected to PowerPoint' : 'Not connected'}</span>
                        ${!this.isConnected ? 
                            '<button class="btn btn-sm btn-primary" onclick="ppt.connectToPowerPoint()">Connect</button>' : 
                            '<button class="btn btn-sm btn-secondary" onclick="ppt.disconnect()">Disconnect</button>'}
                    </div>
                    
                    <div class="export-options-section">
                        <h4>Export Options</h4>
                        
                        <div class="options-grid">
                            <div class="option-group">
                                <label>Slide Layout</label>
                                <select id="slideLayout">
                                    <option value="content" ${this.exportSettings.slideLayout === 'content' ? 'selected' : ''}>Content Slides</option>
                                    <option value="title" ${this.exportSettings.slideLayout === 'title' ? 'selected' : ''}>Title Focus</option>
                                    <option value="comparison" ${this.exportSettings.slideLayout === 'comparison' ? 'selected' : ''}>Comparison</option>
                                    <option value="report" ${this.exportSettings.slideLayout === 'report' ? 'selected' : ''}>Report Format</option>
                                </select>
                            </div>
                            
                            <div class="option-group">
                                <label>Include VBA Macro</label>
                                <select id="includeVBA">
                                    <option value="true" ${this.exportSettings.includeVBA ? 'selected' : ''}>Yes</option>
                                    <option value="false" ${!this.exportSettings.includeVBA ? 'selected' : ''}>No</option>
                                </select>
                            </div>
                            
                            <div class="option-group">
                                <label>Export Theme</label>
                                <select id="theme">
                                    <option value="flowsim" ${this.exportSettings.theme === 'flowsim' ? 'selected' : ''}>FlowSim Blue</option>
                                    <option value="professional" ${this.exportSettings.theme === 'professional' ? 'selected' : ''}>Professional</option>
                                    <option value="technical" ${this.exportSettings.theme === 'technical' ? 'selected' : ''}>Technical</option>
                                    <option value="educational" ${this.exportSettings.theme === 'educational' ? 'selected' : ''}>Educational</option>
                                </select>
                            </div>
                            
                            <div class="option-group">
                                <label>Auto-open after export</label>
                                <select id="autoOpen">
                                    <option value="true" ${this.exportSettings.autoOpen ? 'selected' : ''}>Yes</option>
                                    <option value="false" ${!this.exportSettings.autoOpen ? 'selected' : ''}>No</option>
                                </select>
                            </div>
                        </div>
                        
                        <div class="quick-export-buttons">
                            <button class="btn btn-primary" onclick="ppt.exportToPowerPoint({exportFormat: 'quick'})">
                                <i class="fas fa-bolt"></i> Quick Export
                            </button>
                            <button class="btn btn-secondary" onclick="ppt.exportToPowerPoint({exportFormat: 'full'})">
                                <i class="fas fa-file-alt"></i> Full Report
                            </button>
                            <button class="btn btn-success" onclick="ppt.exportToPowerPoint({exportFormat: 'presentation'})">
                                <i class="fas fa-presentation"></i> Presentation Mode
                            </button>
                        </div>
                    </div>
                    
                    <div class="vba-preview">
                        <h4>VBA Macro Preview</h4>
                        <div class="vba-code">
                            ${this.generateVBAMacroPreview()}
                        </div>
                        <div class="vba-actions">
                            <button class="btn btn-sm btn-secondary" data-action="copy-vba">
                                <i class="fas fa-copy"></i> Copy Code
                            </button>
                            <button class="btn btn-sm btn-primary" data-action="download-vba">
                                <i class="fas fa-download"></i> Download .bas
                            </button>
                        </div>
                    </div>
                </div>
                
                <div class="modal-footer">
                    <button class="btn btn-secondary" onclick="this.closest('.modal').remove()">
                        Cancel
                    </button>
                    <button class="btn btn-primary" onclick="ppt.exportToPowerPoint()">
                        <i class="fas fa-file-powerpoint"></i> Export Now
                    </button>
                </div>
            </div>
        `;
        
        return modal;
    }
    
    generateVBAMacroPreview() {
        const material = window.appState?.currentMaterial;
        if (!material) return '<em>No material selected</em>';
        
        const preview = this.vbaTemplates.basic
            .replace(/{MATERIAL_NAME}/g, material.name || 'Unknown Material')
            .replace(/{DENSITY}/g, material.properties?.density?.toFixed(2) || '0.00')
            .replace(/{FRICTION}/g, material.properties?.friction?.toFixed(2) || '0.00')
            .replace(/{ELASTICITY}/g, material.properties?.elasticity?.toFixed(2) || '0.00');
        
        // Return first 10 lines for preview
        return preview.split('\n').slice(0, 15).join('\n') + '\n...';
    }
    
    copyVBACode() {
        const material = window.appState?.currentMaterial;
        if (!material) {
            showNotification('Please select a material first', 'warning');
            return;
        }
        
        const vbaCode = this.generateVBAMacro({
            materials: { current: material },
            metadata: { userMode: window.appState?.userMode || 'engineer' }
        });
        
        navigator.clipboard.writeText(vbaCode).then(() => {
            showNotification('VBA code copied to clipboard!', 'success');
        }).catch(err => {
            console.error('Copy failed:', err);
            showNotification('Failed to copy code', 'error');
        });
    }
    
    downloadVBAMacro() {
        const material = window.appState?.currentMaterial;
        if (!material) {
            showNotification('Please select a material first', 'warning');
            return;
        }
        
        const vbaCode = this.generateVBAMacro({
            materials: { current: material },
            metadata: { userMode: window.appState?.userMode || 'engineer' }
        });
        
        const blob = new Blob([vbaCode], { type: 'text/plain' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `FlowSim_${material.name.replace(/\s+/g, '_')}_Macro.bas`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        
        showNotification('VBA macro downloaded!', 'success');
    }
    
    // ====== DEFAULT TEMPLATES ======
    
    getDefaultTemplates() {
        return {
            title: {
                name: 'Title Slide',
                layout: 'title',
                placeholders: ['title', 'subtitle', 'date', 'author'],
                theme: 'flowsim_title'
            },
            content: {
                name: 'Content Slide',
                layout: 'twoColumn',
                placeholders: ['title', 'leftContent', 'rightContent'],
                theme: 'flowsim_content'
            },
            analysis: {
                name: 'Analysis Slide',
                layout: 'full',
                placeholders: ['title', 'charts', 'analysis'],
                theme: 'flowsim_analysis'
            },
            comparison: {
                name: 'Comparison Slide',
                layout: 'comparison',
                placeholders: ['title', 'table', 'notes'],
                theme: 'flowsim_comparison'
            }
        };
    }
    
    async fetchTemplates() {
        // In production, fetch from API
        // For demo, return default templates
        return this.getDefaultTemplates();
    }
    
    // ====== LIVE PRESENTATION CONTROL ======
    
    async startLivePresentation() {
        if (!this.isConnected) {
            const connected = await this.connectToPowerPoint();
            if (!connected) return false;
        }
        
        showNotification('Starting live presentation control...', 'info');
        
        this.liveControl = {
            active: true,
            startTime: new Date(),
            slideIndex: 1,
            participants: 0,
            controls: {
                remoteControl: true,
                audienceInteraction: true,
                liveAnnotations: true
            }
        };
        
        // Generate QR code for audience joining
        this.generateAudienceQR();
        
        // Start WebSocket connection for live updates
        this.startLiveWebSocket();
        
        showNotification('Live presentation started! Audience can join via QR code.', 'success');
        
        return true;
    }
    
    generateAudienceQR() {
        // Generate QR code for audience to join presentation
        // This would use a QR code library in production
        const qrContainer = document.querySelector('.qr-container');
        if (qrContainer) {
            qrContainer.innerHTML = `
                <div class="qr-code">
                    <div class="qr-placeholder">
                        <i class="fas fa-qrcode"></i>
                        <p>Scan to join live presentation</p>
                        <small>Code: FS${Date.now().toString(36).toUpperCase()}</small>
                    </div>
                </div>
            `;
        }
    }
    
    startLiveWebSocket() {
        // Start WebSocket connection for live presentation control
        // This is a simulation for demo purposes
        console.log('📡 Starting live presentation WebSocket (simulated)');
        
        // Simulate audience joining
        setInterval(() => {
            if (this.liveControl && this.liveControl.active) {
                this.liveControl.participants += Math.floor(Math.random() * 3);
                this.updateLiveParticipantCount();
            }
        }, 5000);
    }
    
    updateLiveParticipantCount() {
        const countElement = document.querySelector('.live-participants');
        if (countElement && this.liveControl) {
            countElement.textContent = this.liveControl.participants;
        }
    }
    
    // ====== PUBLIC API ======
    
    getExportSettings() {
        return { ...this.exportSettings };
    }
    
    setExportSetting(key, value) {
        this.exportSettings[key] = value;
        this.saveSettings();
    }
    
    getVBATemplate(type = 'basic') {
        return this.vbaTemplates[type] || this.vbaTemplates.basic;
    }
    
    generateCustomVBA(templateType, data) {
        const template = this.getVBATemplate(templateType);
        
        // Replace placeholders in template
        let vbaCode = template;
        for (const [key, value] of Object.entries(data)) {
            const placeholder = `{${key.toUpperCase()}}`;
            vbaCode = vbaCode.replace(new RegExp(placeholder, 'g'), value);
        }
        
        return vbaCode;
    }
    
    // ====== STATIC METHODS ======
    
    static checkOfficeAvailability() {
        return typeof Office !== 'undefined' && Office.context;
    }
    
    static getSupportedFormats() {
        return ['pptx', 'pdf', 'png', 'jpg'];
    }
}

// ====== GLOBAL PPT INTEGRATION INSTANCE ======

let ppt = null;

// Initialize when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    ppt = new PowerPointIntegration();
    console.log('📊 PowerPoint Integration ready');
    
    // Make globally available
    window.ppt = ppt;
    
    // Add event listeners for UI buttons
    const pptButtons = document.querySelectorAll('[data-action="export-ppt"]');
    pptButtons.forEach(btn => {
        btn.addEventListener('click', () => ppt.exportToPowerPoint());
    });
});

// ====== PPT INTEGRATION STYLES ======

const pptStyles = document.createElement('style');
pptStyles.textContent = `
    .export-modal {
        max-width: 800px;
        max-height: 90vh;
        overflow-y: auto;
    }
    
    .modal-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 20px;
        background: linear-gradient(135deg, #ff7b00 0%, #d44500 100%);
        color: white;
    }
    
    .modal-header h3 {
        margin: 0;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    
    .modal-body {
        padding: 20px;
    }
    
    .connection-status {
        display: flex;
        align-items: center;
        gap: 10px;
        padding: 15px;
        background: var(--gray-lighter);
        border-radius: var(--radius-md);
        margin-bottom: 20px;
    }
    
    .export-options-section {
        margin-bottom: 30px;
    }
    
    .export-options-section h4 {
        margin-bottom: 15px;
        color: var(--dark);
    }
    
    .options-grid {
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 20px;
        margin-bottom: 20px;
    }
    
    .option-group {
        display: flex;
        flex-direction: column;
        gap: 8px;
    }
    
    .option-group label {
        font-size: 14px;
        font-weight: 600;
        color: var(--dark);
    }
    
    .option-group select {
        padding: 8px 12px;
        border: 1px solid var(--gray-light);
        border-radius: var(--radius-md);
        background: white;
        color: var(--dark);
        font-size: 14px;
        cursor: pointer;
    }
    
    .quick-export-buttons {
        display: flex;
        gap: 10px;
        margin-top: 20px;
        flex-wrap: wrap;
    }
    
    .vba-preview {
        background: var(--gray-lighter);
        border-radius: var(--radius-md);
        padding: 20px;
        margin: 20px 0;
    }
    
    .vba-preview h4 {
        margin-bottom: 15px;
        color: var(--dark);
    }
    
    .vba-code {
        background: var(--dark);
        color: #00ff00;
        font-family: 'Consolas', 'Monaco', monospace;
        font-size: 12px;
        padding: 15px;
        border-radius: var(--radius-sm);
        overflow-x: auto;
        white-space: pre-wrap;
        max-height: 200px;
        overflow-y: auto;
        margin-bottom: 15px;
    }
    
    .vba-actions {
        display: flex;
        gap: 10px;
        justify-content: flex-end;
    }
    
    .modal-footer {
        display: flex;
        justify-content: flex-end;
        gap: 10px;
        padding: 20px;
        border-top: 1px solid var(--gray-light);
    }
    
    .qr-code {
        text-align: center;
        padding: 20px;
    }
    
    .qr-placeholder {
        background: white;
        padding: 30px;
        border-radius: var(--radius-md);
        display: inline-block;
        box-shadow: var(--shadow-md);
    }
    
    .qr-placeholder i {
        font-size: 60px;
        color: var(--primary);
        margin-bottom: 15px;
    }
    
    .qr-placeholder p {
        margin: 10px 0 5px 0;
        font-weight: 600;
        color: var(--dark);
    }
    
    .qr-placeholder small {
        color: var(--gray);
        font-size: 12px;
    }
    
    .live-controls {
        display: flex;
        gap: 10px;
        align-items: center;
        margin-top: 20px;
        padding: 15px;
        background: var(--gray-lighter);
        border-radius: var(--radius-md);
    }
    
    .participant-count {
        display: flex;
        align-items: center;
        gap: 8px;
        background: white;
        padding: 8px 15px;
        border-radius: var(--radius-md);
        font-weight: 600;
    }
    
    .participant-count i {
        color: var(--primary);
    }
`;

document.head.appendChild(pptStyles);

console.log('📊 PowerPoint Integration module loaded');
