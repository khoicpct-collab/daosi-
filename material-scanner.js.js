// ====== FLOWSIM MATERIAL STUDIO - AI MATERIAL SCANNER ======

class MaterialScannerAI {
    constructor(apiEndpoint = null) {
        this.apiEndpoint = apiEndpoint;
        this.model = null;
        this.isModelLoaded = false;
        this.cameraStream = null;
        this.scanningActive = false;
        this.scanResults = [];
        this.scanCanvas = null;
        this.scanContext = null;
        
        // Material database for classification
        this.materialDatabase = {
            metals: {
                steel: { density: 7.85, friction: 0.3, elasticity: 0.8, color: '#808080' },
                aluminum: { density: 2.7, friction: 0.4, elasticity: 0.9, color: '#C0C0C0' },
                copper: { density: 8.96, friction: 0.35, elasticity: 0.7, color: '#B87333' },
                brass: { density: 8.5, friction: 0.4, elasticity: 0.75, color: '#B5A642' }
            },
            plastics: {
                pvc: { density: 1.38, friction: 0.4, elasticity: 0.3, color: '#F0F8FF' },
                pe: { density: 0.94, friction: 0.3, elasticity: 0.5, color: '#FFE4E1' },
                pp: { density: 0.9, friction: 0.25, elasticity: 0.6, color: '#F0FFFF' },
                abs: { density: 1.05, friction: 0.45, elasticity: 0.4, color: '#FFE4C4' }
            },
            granular: {
                sand: { density: 1.6, friction: 0.7, elasticity: 0.2, color: '#D4A76A' },
                gravel: { density: 1.68, friction: 0.65, elasticity: 0.15, color: '#8B7355' },
                rice: { density: 0.85, friction: 0.6, elasticity: 0.25, color: '#FAFAD2' },
                coffee: { density: 0.56, friction: 0.5, elasticity: 0.1, color: '#8B4513' }
            },
            organic: {
                wood: { density: 0.6, friction: 0.5, elasticity: 0.4, color: '#8B4513' },
                grain: { density: 0.75, friction: 0.6, elasticity: 0.3, color: '#F4D03F' },
                biomass: { density: 0.45, friction: 0.55, elasticity: 0.2, color: '#228B22' }
            }
        };
        
        // QR/Barcode detection
        this.barcodeDetector = null;
        this.initBarcodeDetector();
        
        console.log('🤖 AI Material Scanner initialized');
    }
    
    // ====== INITIALIZATION ======
    
    async init() {
        try {
            // Load TensorFlow.js model (if available)
            await this.loadModel();
            
            // Initialize camera access
            await this.initCamera();
            
            // Initialize canvas for processing
            this.initCanvas();
            
            console.log('✅ Material Scanner ready');
            return true;
            
        } catch (error) {
            console.error('Scanner initialization failed:', error);
            return false;
        }
    }
    
    async initBarcodeDetector() {
        if ('BarcodeDetector' in window) {
            try {
                const supportedFormats = await BarcodeDetector.getSupportedFormats();
                this.barcodeDetector = new BarcodeDetector({ formats: supportedFormats });
                console.log('📷 BarcodeDetector initialized with formats:', supportedFormats);
            } catch (error) {
                console.warn('BarcodeDetector not available:', error);
            }
        } else {
            console.warn('BarcodeDetector API not supported in this browser');
        }
    }
    
    async loadModel() {
        // Check if TensorFlow.js is available
        if (typeof tf === 'undefined') {
            console.warn('TensorFlow.js not loaded, using fallback analysis');
            this.isModelLoaded = false;
            return;
        }
        
        try {
            // In production, load a pre-trained model
            // For demo, we'll simulate model loading
            console.log('🧠 Loading AI model...');
            
            // Simulate model loading delay
            await new Promise(resolve => setTimeout(resolve, 1000));
            
            // Create a simple model for demonstration
            this.model = {
                predict: (imageData) => this.simulatePrediction(imageData)
            };
            
            this.isModelLoaded = true;
            console.log('✅ AI model loaded (simulated)');
            
        } catch (error) {
            console.error('Failed to load AI model:', error);
            this.isModelLoaded = false;
        }
    }
    
    // ====== CAMERA HANDLING ======
    
    async initCamera() {
        try {
            // Request camera permission
            const stream = await navigator.mediaDevices.getUserMedia({
                video: {
                    facingMode: 'environment', // Prefer rear camera
                    width: { ideal: 1280 },
                    height: { ideal: 720 }
                },
                audio: false
            });
            
            this.cameraStream = stream;
            console.log('📷 Camera access granted');
            return stream;
            
        } catch (error) {
            console.error('Camera access denied:', error);
            throw new Error('Camera access required for material scanning');
        }
    }
    
    initCanvas() {
        this.scanCanvas = document.createElement('canvas');
        this.scanContext = this.scanCanvas.getContext('2d', { willReadFrequently: true });
        console.log('🎨 Processing canvas initialized');
    }
    
    // ====== SCANNING FUNCTIONS ======
    
    async startScanning(videoElement, canvasElement, onResult = null) {
        if (!this.cameraStream) {
            throw new Error('Camera not initialized');
        }
        
        this.scanningActive = true;
        const video = videoElement;
        const canvas = canvasElement;
        const ctx = canvas.getContext('2d');
        
        // Set video source
        video.srcObject = this.cameraStream;
        await video.play();
        
        console.log('🔍 Starting material scan...');
        
        // Scanning loop
        const scanLoop = async () => {
            if (!this.scanningActive) return;
            
            try {
                // Draw video frame to canvas
                canvas.width = video.videoWidth;
                canvas.height = video.videoHeight;
                ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
                
                // Process frame for scanning
                const scanFrame = await this.processFrame(canvas);
                
                // If we found something, trigger callback
                if (scanFrame && onResult) {
                    onResult(scanFrame);
                }
                
                // Continue scanning
                requestAnimationFrame(scanLoop);
                
            } catch (error) {
                console.error('Scanning error:', error);
                this.stopScanning();
            }
        };
        
        scanLoop();
    }
    
    stopScanning() {
        this.scanningActive = false;
        
        // Stop camera stream
        if (this.cameraStream) {
            this.cameraStream.getTracks().forEach(track => track.stop());
            this.cameraStream = null;
        }
        
        console.log('🛑 Scanning stopped');
    }
    
    // ====== FRAME PROCESSING ======
    
    async processFrame(canvas) {
        const ctx = canvas.getContext('2d');
        const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
        
        let results = {
            timestamp: Date.now(),
            materials: [],
            barcodes: [],
            imageData: imageData,
            frameSize: { width: canvas.width, height: canvas.height }
        };
        
        // 1. Check for barcodes/QR codes first
        if (this.barcodeDetector) {
            try {
                const barcodes = await this.barcodeDetector.detect(canvas);
                if (barcodes.length > 0) {
                    results.barcodes = barcodes;
                    await this.processBarcodes(barcodes, results);
                }
            } catch (error) {
                console.warn('Barcode detection failed:', error);
            }
        }
        
        // 2. Analyze material if no barcode found
        if (results.barcodes.length === 0) {
            await this.analyzeMaterial(imageData, results);
        }
        
        // 3. Store result
        this.scanResults.push(results);
        if (this.scanResults.length > 50) {
            this.scanResults.shift(); // Keep only last 50 results
        }
        
        return results;
    }
    
    async processBarcodes(barcodes, results) {
        for (const barcode of barcodes) {
            const barcodeData = {
                rawValue: barcode.rawValue,
                format: barcode.format,
                boundingBox: barcode.boundingBox,
                cornerPoints: barcode.cornerPoints
            };
            
            console.log('📊 Barcode detected:', barcodeData);
            
            // Try to interpret barcode as material code
            const materialInfo = this.interpretMaterialCode(barcode.rawValue);
            if (materialInfo) {
                results.materials.push(materialInfo);
                results.source = 'barcode';
                results.confidence = 0.95;
            }
        }
    }
    
    interpretMaterialCode(code) {
        // Simple material code interpretation
        // In production, this would query a database
        const codeMap = {
            'MAT001': { name: 'Steel Ball', category: 'metals', type: 'steel' },
            'MAT002': { name: 'Plastic Pellet', category: 'plastics', type: 'pe' },
            'MAT003': { name: 'Sand', category: 'granular', type: 'sand' },
            'MAT004': { name: 'Wood Chip', category: 'organic', type: 'wood' },
            'MAT005': { name: 'Grain', category: 'agricultural', type: 'grain' }
        };
        
        if (codeMap[code]) {
            const baseInfo = codeMap[code];
            const properties = this.materialDatabase[baseInfo.category]?.[baseInfo.type];
            
            if (properties) {
                return {
                    id: `scanned_${Date.now()}`,
                    name: baseInfo.name,
                    category: baseInfo.category,
                    type: baseInfo.type,
                    properties: { ...properties },
                    tags: [baseInfo.category, baseInfo.type, 'scanned'],
                    aiConfidence: 0.95,
                    source: 'barcode',
                    code: code
                };
            }
        }
        
        return null;
    }
    
    // ====== MATERIAL ANALYSIS ======
    
    async analyzeMaterial(imageData, results) {
        try {
            let analysisResult;
            
            if (this.isModelLoaded && this.model) {
                // Use AI model for analysis
                analysisResult = await this.model.predict(imageData);
            } else {
                // Fallback to computer vision analysis
                analysisResult = await this.computerVisionAnalysis(imageData);
            }
            
            if (analysisResult) {
                results.materials.push(analysisResult);
                results.source = this.isModelLoaded ? 'ai_model' : 'computer_vision';
                results.confidence = analysisResult.aiConfidence || 0.7;
            }
            
        } catch (error) {
            console.error('Material analysis failed:', error);
            // Provide fallback result
            results.materials.push(this.getFallbackMaterial());
            results.source = 'fallback';
            results.confidence = 0.5;
        }
    }
    
    simulatePrediction(imageData) {
        // Simulate AI prediction
        // In production, this would run actual TensorFlow.js model
        
        // Analyze image colors
        const colorAnalysis = this.analyzeColors(imageData);
        const textureAnalysis = this.analyzeTexture(imageData);
        
        // Determine material type based on analysis
        const materialType = this.classifyMaterial(colorAnalysis, textureAnalysis);
        const properties = this.estimateProperties(materialType, colorAnalysis, textureAnalysis);
        
        return {
            id: `scanned_${Date.now()}`,
            name: this.getMaterialName(materialType),
            category: this.getMaterialCategory(materialType),
            type: materialType,
            properties: properties,
            tags: this.generateTags(materialType, colorAnalysis),
            aiConfidence: this.calculateConfidence(colorAnalysis, textureAnalysis),
            analysisData: {
                colors: colorAnalysis,
                texture: textureAnalysis
            }
        };
    }
    
    async computerVisionAnalysis(imageData) {
        // Computer vision based material analysis
        const colorAnalysis = this.analyzeColors(imageData);
        const textureAnalysis = this.analyzeTexture(imageData);
        const edgeAnalysis = this.analyzeEdges(imageData);
        
        const materialType = this.classifyMaterialCV(colorAnalysis, textureAnalysis, edgeAnalysis);
        const properties = this.estimateProperties(materialType, colorAnalysis, textureAnalysis);
        
        return {
            id: `scanned_${Date.now()}`,
            name: this.getMaterialName(materialType),
            category: this.getMaterialCategory(materialType),
            type: materialType,
            properties: properties,
            tags: this.generateTags(materialType, colorAnalysis),
            aiConfidence: this.calculateCVConfidence(colorAnalysis, textureAnalysis, edgeAnalysis),
            analysisData: {
                colors: colorAnalysis,
                texture: textureAnalysis,
                edges: edgeAnalysis
            }
        };
    }
    
    // ====== IMAGE ANALYSIS FUNCTIONS ======
    
    analyzeColors(imageData) {
        const data = imageData.data;
        const pixelCount = data.length / 4;
        
        let r = 0, g = 0, b = 0;
        let rMin = 255, gMin = 255, bMin = 255;
        let rMax = 0, gMax = 0, bMax = 0;
        
        // Calculate color statistics
        for (let i = 0; i < data.length; i += 4) {
            r += data[i];
            g += data[i + 1];
            b += data[i + 2];
            
            rMin = Math.min(rMin, data[i]);
            gMin = Math.min(gMin, data[i + 1]);
            bMin = Math.min(bMin, data[i + 2]);
            
            rMax = Math.max(rMax, data[i]);
            gMax = Math.max(gMax, data[i + 1]);
            bMax = Math.max(bMax, data[i + 2]);
        }
        
        const avgR = Math.round(r / pixelCount);
        const avgG = Math.round(g / pixelCount);
        const avgB = Math.round(b / pixelCount);
        
        // Calculate color variance
        let variance = 0;
        for (let i = 0; i < data.length; i += 4) {
            const dr = data[i] - avgR;
            const dg = data[i + 1] - avgG;
            const db = data[i + 2] - avgB;
            variance += (dr * dr + dg * dg + db * db);
        }
        variance /= pixelCount;
        
        return {
            average: { r: avgR, g: avgG, b: avgB },
            range: {
                r: { min: rMin, max: rMax },
                g: { min: gMin, max: gMax },
                b: { min: bMin, max: bMax }
            },
            variance: variance,
            dominantColor: this.rgbToHex(avgR, avgG, avgB),
            isMonochrome: variance < 1000,
            brightness: (avgR + avgG + avgB) / 3
        };
    }
    
    analyzeTexture(imageData) {
        const width = imageData.width;
        const height = imageData.height;
        const data = imageData.data;
        
        let contrast = 0;
        let uniformity = 0;
        let edgeCount = 0;
        
        // Simple texture analysis using edge detection
        for (let y = 1; y < height - 1; y++) {
            for (let x = 1; x < width - 1; x++) {
                const idx = (y * width + x) * 4;
                
                // Sobel operator for edge detection
                const gx = (
                    -this.getGray(data, (y-1)*width*4 + (x-1)*4) + this.getGray(data, (y-1)*width*4 + (x+1)*4) +
                    -2 * this.getGray(data, y*width*4 + (x-1)*4) + 2 * this.getGray(data, y*width*4 + (x+1)*4) +
                    -this.getGray(data, (y+1)*width*4 + (x-1)*4) + this.getGray(data, (y+1)*width*4 + (x+1)*4)
                );
                
                const gy = (
                    -this.getGray(data, (y-1)*width*4 + (x-1)*4) - 2 * this.getGray(data, (y-1)*width*4 + x*4) - this.getGray(data, (y-1)*width*4 + (x+1)*4) +
                    this.getGray(data, (y+1)*width*4 + (x-1)*4) + 2 * this.getGray(data, (y+1)*width*4 + x*4) + this.getGray(data, (y+1)*width*4 + (x+1)*4)
                );
                
                const gradient = Math.sqrt(gx * gx + gy * gy);
                
                if (gradient > 50) { // Edge threshold
                    edgeCount++;
                }
                
                contrast += gradient;
            }
        }
        
        const totalPixels = width * height;
        const edgeDensity = edgeCount / totalPixels;
        
        return {
            contrast: contrast / totalPixels,
            edgeDensity: edgeDensity,
            textureType: this.classifyTexture(edgeDensity),
            granularity: edgeDensity > 0.1 ? 'fine' : edgeDensity > 0.05 ? 'medium' : 'coarse'
        };
    }
    
    analyzeEdges(imageData) {
        // Simplified edge analysis
        const width = imageData.width;
        const height = imageData.height;
        const data = imageData.data;
        
        let horizontalEdges = 0;
        let verticalEdges = 0;
        
        for (let y = 0; y < height; y++) {
            for (let x = 0; x < width; x++) {
                const idx = (y * width + x) * 4;
                
                if (x < width - 1) {
                    const nextIdx = (y * width + (x + 1)) * 4;
                    const diff = Math.abs(data[idx] - data[nextIdx]) + 
                                 Math.abs(data[idx + 1] - data[nextIdx + 1]) + 
                                 Math.abs(data[idx + 2] - data[nextIdx + 2]);
                    if (diff > 30) horizontalEdges++;
                }
                
                if (y < height - 1) {
                    const nextIdx = ((y + 1) * width + x) * 4;
                    const diff = Math.abs(data[idx] - data[nextIdx]) + 
                                 Math.abs(data[idx + 1] - data[nextIdx + 1]) + 
                                 Math.abs(data[idx + 2] - data[nextIdx + 2]);
                    if (diff > 30) verticalEdges++;
                }
            }
        }
        
        return {
            horizontal: horizontalEdges,
            vertical: verticalEdges,
            total: horizontalEdges + verticalEdges,
            orientation: horizontalEdges > verticalEdges * 1.5 ? 'horizontal' : 
                        verticalEdges > horizontalEdges * 1.5 ? 'vertical' : 'mixed'
        };
    }
    
    getGray(data, index) {
        return data[index] * 0.299 + data[index + 1] * 0.587 + data[index + 2] * 0.114;
    }
    
    // ====== MATERIAL CLASSIFICATION ======
    
    classifyMaterial(colorAnalysis, textureAnalysis) {
        const { average, brightness, variance } = colorAnalysis;
        const { edgeDensity, granularity } = textureAnalysis;
        
        // Simple classification rules
        if (brightness > 200 && variance < 500) {
            return 'plastic_light';
        } else if (brightness < 100 && variance < 300) {
            return 'metal_dark';
        } else if (edgeDensity > 0.08 && granularity === 'fine') {
            return 'sand';
        } else if (brightness > 150 && brightness < 180 && variance > 1000) {
            return 'wood';
        } else if (average.b > average.r && average.b > average.g) {
            return 'plastic_blue';
        } else if (average.r > average.g * 1.5 && average.r > average.b * 1.5) {
            return 'copper';
        } else {
            return 'unknown';
        }
    }
    
    classifyMaterialCV(colorAnalysis, textureAnalysis, edgeAnalysis) {
        // More sophisticated classification using computer vision
        const { brightness, variance, isMonochrome } = colorAnalysis;
        const { edgeDensity, textureType } = textureAnalysis;
        const { orientation, total } = edgeAnalysis;
        
        if (isMonochrome && brightness > 180) {
            return 'aluminum';
        } else if (isMonochrome && brightness < 100) {
            return 'steel';
        } else if (edgeDensity > 0.1 && textureType === 'granular') {
            return 'granular_fine';
        } else if (variance > 1500 && textureType === 'irregular') {
            return 'organic';
        } else if (orientation === 'horizontal' && total > 10000) {
            return 'textured';
        } else {
            return 'composite';
        }
    }
    
    classifyTexture(edgeDensity) {
        if (edgeDensity < 0.02) return 'smooth';
        if (edgeDensity < 0.05) return 'slightly_textured';
        if (edgeDensity < 0.1) return 'textured';
        if (edgeDensity < 0.2) return 'granular';
        return 'irregular';
    }
    
    // ====== PROPERTY ESTIMATION ======
    
    estimateProperties(materialType, colorAnalysis, textureAnalysis) {
        const baseProperties = {
            density: 1.0,
            friction: 0.5,
            elasticity: 0.5,
            size: 'medium',
            color: colorAnalysis.dominantColor
        };
        
        // Adjust properties based on material type
        switch (materialType) {
            case 'steel':
            case 'metal_dark':
                baseProperties.density = 7.8;
                baseProperties.friction = 0.3;
                baseProperties.elasticity = 0.8;
                baseProperties.size = 'medium';
                break;
                
            case 'aluminum':
                baseProperties.density = 2.7;
                baseProperties.friction = 0.4;
                baseProperties.elasticity = 0.9;
                baseProperties.size = 'medium';
                break;
                
            case 'plastic_light':
            case 'plastic_blue':
                baseProperties.density = 1.2;
                baseProperties.friction = 0.4;
                baseProperties.elasticity = 0.6;
                baseProperties.size = 'small';
                break;
                
            case 'sand':
            case 'granular_fine':
                baseProperties.density = 1.6;
                baseProperties.friction = 0.7;
                baseProperties.elasticity = 0.2;
                baseProperties.size = 'fine';
                break;
                
            case 'wood':
            case 'organic':
                baseProperties.density = 0.6;
                baseProperties.friction = 0.5;
                baseProperties.elasticity = 0.4;
                baseProperties.size = 'medium';
                break;
                
            case 'copper':
                baseProperties.density = 8.9;
                baseProperties.friction = 0.35;
                baseProperties.elasticity = 0.7;
                baseProperties.size = 'medium';
                baseProperties.color = '#B87333';
                break;
                
            default:
                // Use texture analysis to estimate
                baseProperties.density = 1.0 + textureAnalysis.edgeDensity * 5;
                baseProperties.friction = 0.3 + textureAnalysis.edgeDensity * 2;
                baseProperties.elasticity = 0.7 - textureAnalysis.edgeDensity * 3;
        }
        
        return baseProperties;
    }
    
    // ====== UTILITY FUNCTIONS ======
    
    getMaterialName(type) {
        const nameMap = {
            'steel': 'Steel',
            'metal_dark': 'Dark Metal',
            'aluminum': 'Aluminum',
            'plastic_light': 'Light Plastic',
            'plastic_blue': 'Blue Plastic',
            'sand': 'Sand',
            'granular_fine': 'Fine Granular Material',
            'wood': 'Wood',
            'organic': 'Organic Material',
            'copper': 'Copper',
            'textured': 'Textured Material',
            'composite': 'Composite Material',
            'unknown': 'Unknown Material'
        };
        
        return nameMap[type] || 'Scanned Material';
    }
    
    getMaterialCategory(type) {
        if (type.includes('metal') || type === 'steel' || type === 'aluminum' || type === 'copper') {
            return 'metals';
        } else if (type.includes('plastic')) {
            return 'plastics';
        } else if (type.includes('granular') || type === 'sand') {
            return 'granular';
        } else if (type.includes('wood') || type === 'organic') {
            return 'organic';
        } else {
            return 'composite';
        }
    }
    
    generateTags(materialType, colorAnalysis) {
        const tags = [materialType];
        
        // Add color tags
        if (colorAnalysis.brightness > 180) tags.push('light');
        if (colorAnalysis.brightness < 100) tags.push('dark');
        if (colorAnalysis.isMonochrome) tags.push('monochrome');
        else tags.push('colorful');
        
        // Add texture tags
        if (colorAnalysis.variance > 1000) tags.push('textured');
        
        return tags.slice(0, 5); // Limit to 5 tags
    }
    
    calculateConfidence(colorAnalysis, textureAnalysis) {
        let confidence = 0.7; // Base confidence
        
        // Increase confidence for clear characteristics
        if (colorAnalysis.isMonochrome) confidence += 0.1;
        if (colorAnalysis.variance < 500 || colorAnalysis.variance > 1500) confidence += 0.1;
        if (textureAnalysis.edgeDensity > 0.1) confidence += 0.1;
        
        return Math.min(0.95, confidence);
    }
    
    calculateCVConfidence(colorAnalysis, textureAnalysis, edgeAnalysis) {
        let confidence = 0.6;
        
        // Multiple analysis points increase confidence
        const analysisPoints = [
            colorAnalysis.isMonochrome,
            colorAnalysis.variance > 1000,
            textureAnalysis.edgeDensity > 0.05,
            edgeAnalysis.total > 5000
        ];
        
        const validPoints = analysisPoints.filter(Boolean).length;
        confidence += validPoints * 0.08;
        
        return Math.min(0.92, confidence);
    }
    
    getFallbackMaterial() {
        return {
            id: `fallback_${Date.now()}`,
            name: 'Unknown Material',
            category: 'unknown',
            type: 'unknown',
            properties: {
                density: 1.0,
                friction: 0.5,
                elasticity: 0.5,
                size: 'medium',
                color: '#808080'
            },
            tags: ['unknown', 'scanned'],
            aiConfidence: 0.5,
            source: 'fallback'
        };
    }
    
    // ====== IMAGE UPLOAD PROCESSING ======
    
    async processImageFile(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            
            reader.onload = async (event) => {
                try {
                    const img = new Image();
                    
                    img.onload = async () => {
                        // Create canvas for processing
                        const canvas = document.createElement('canvas');
                        const ctx = canvas.getContext('2d');
                        
                        canvas.width = img.width;
                        canvas.height = img.height;
                        ctx.drawImage(img, 0, 0);
                        
                        // Process the image
                        const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
                        const result = await this.processFrame(canvas);
                        
                        resolve(result);
                    };
                    
                    img.onerror = () => reject(new Error('Failed to load image'));
                    img.src = event.target.result;
                    
                } catch (error) {
                    reject(error);
                }
            };
            
            reader.onerror = () => reject(new Error('Failed to read file'));
            reader.readAsDataURL(file);
        });
    }
    
    // ====== SCANNER UI INTEGRATION ======
    
    async openScannerModal() {
        // Create scanner modal
        const modal = this.createScannerModal();
        document.body.appendChild(modal);
        
        // Initialize scanner
        await this.init();
        
        // Start scanning
        const video = modal.querySelector('.camera-feed');
        const canvas = modal.querySelector('.camera-canvas');
        
        this.startScanning(video, canvas, (result) => {
            this.updateScannerUI(modal, result);
        });
        
        return modal;
    }
    
    createScannerModal() {
        const modal = document.createElement('div');
        modal.className = 'modal visible';
        modal.innerHTML = `
            <div class="modal-content scanner-modal">
                <div class="scanner-header">
                    <h3><i class="fas fa-camera"></i> Material Scanner</h3>
                    <button class="btn-close" onclick="this.closest('.modal').remove(); scanner.stopScanning()">
                        <i class="fas fa-times"></i>
                    </button>
                </div>
                
                <div class="scanner-container">
                    <div class="camera-section">
                        <div class="camera-header">
                            <span>Live Camera Feed</span>
                            <button class="btn-small btn-toggle-camera">
                                <i class="fas fa-sync-alt"></i> Switch Camera
                            </button>
                        </div>
                        <div class="camera-feed-container">
                            <video class="camera-feed" autoplay playsinline muted></video>
                            <canvas class="camera-canvas"></canvas>
                            <div class="camera-overlay">
                                <div class="scan-frame"></div>
                                <div class="scan-instruction">
                                    <i class="fas fa-arrows-alt-h"></i>
                                    <p>Align material within frame</p>
                                </div>
                            </div>
                        </div>
                        <div class="camera-controls">
                            <button class="btn btn-primary btn-capture">
                                <i class="fas fa-camera"></i> Capture & Analyze
                            </button>
                            <button class="btn btn-secondary btn-toggle-flash">
                                <i class="fas fa-bolt"></i> Flash
                            </button>
                        </div>
                    </div>
                    
                    <div class="analysis-results">
                        <div class="analysis-header">
                            <h4>Analysis Results</h4>
                            <div class="scanner-status">
                                <span class="status-dot"></span>
                                <span>Scanning...</span>
                            </div>
                        </div>
                        
                        <div class="analysis-grid">
                            <div class="analysis-item">
                                <label>Material Name</label>
                                <div class="material-name-display">-</div>
                            </div>
                            <div class="analysis-item">
                                <label>Category</label>
                                <div class="material-category-display">-</div>
                            </div>
                            <div class="analysis-item">
                                <label>Confidence</label>
                                <div class="confidence-display">
                                    <div class="confidence-bar">
                                        <div class="confidence-fill" style="width: 0%"></div>
                                    </div>
                                    <span class="confidence-value">0%</span>
                                </div>
                            </div>
                            <div class="analysis-item">
                                <label>Properties</label>
                                <div class="properties-display">
                                    <span class="property-tag">Density: -</span>
                                    <span class="property-tag">Friction: -</span>
                                    <span class="property-tag">Elasticity: -</span>
                                </div>
                            </div>
                        </div>
                        
                        <div class="ai-tags-container">
                            <label>AI Tags</label>
                            <div class="ai-tags">
                                <!-- Tags will be populated here -->
                            </div>
                        </div>
                        
                        <div class="scanner-actions">
                            <button class="btn btn-secondary btn-rescan">
                                <i class="fas fa-redo"></i> Rescan
                            </button>
                            <button class="btn btn-success btn-use-material" disabled>
                                <i class="fas fa-check"></i> Use This Material
                            </button>
                        </div>
                    </div>
                </div>
            </div>
        `;
        
        // Add event listeners
        modal.querySelector('.btn-capture').addEventListener('click', () => {
            this.captureAndAnalyze(modal);
        });
        
        modal.querySelector('.btn-use-material').addEventListener('click', () => {
            this.useScannedMaterial(modal);
        });
        
        return modal;
    }
    
    updateScannerUI(modal, result) {
        if (!result.materials || result.materials.length === 0) return;
        
        const material = result.materials[0];
        
        // Update UI elements
        modal.querySelector('.material-name-display').textContent = material.name;
        modal.querySelector('.material-category-display').textContent = material.category;
        
        // Update confidence
        const confidence = Math.round((material.aiConfidence || 0.7) * 100);
        modal.querySelector('.confidence-fill').style.width = `${confidence}%`;
        modal.querySelector('.confidence-value').textContent = `${confidence}%`;
        
        // Update properties
        const props = material.properties;
        const propsDisplay = modal.querySelector('.properties-display');
        propsDisplay.innerHTML = `
            <span class="property-tag">Density: ${props.density.toFixed(2)}</span>
            <span class="property-tag">Friction: ${props.friction.toFixed(2)}</span>
            <span class="property-tag">Elasticity: ${props.elasticity.toFixed(2)}</span>
        `;
        
        // Update AI tags
        const tagsContainer = modal.querySelector('.ai-tags');
        tagsContainer.innerHTML = material.tags
            .map(tag => `<span class="ai-tag">${tag}</span>`)
            .join('');
        
        // Enable use button
        modal.querySelector('.btn-use-material').disabled = false;
        
        // Update status
        const statusDot = modal.querySelector('.status-dot');
        statusDot.classList.toggle('active', confidence > 70);
        
        const statusText = modal.querySelector('.scanner-status span:last-child');
        statusText.textContent = confidence > 70 ? 'High Confidence' : 'Low Confidence';
    }
    
    async captureAndAnalyze(modal) {
        const video = modal.querySelector('.camera-feed');
        const canvas = modal.querySelector('.camera-canvas');
        const ctx = canvas.getContext('2d');
        
        // Capture current frame
        canvas.width = video.videoWidth;
        canvas.height = video.videoHeight;
        ctx.drawImage(video, 0, 0);
        
        // Process frame
        const result = await this.processFrame(canvas);
        this.updateScannerUI(modal, result);
        
        // Show confirmation
        this.showScanConfirmation(modal, result.materials[0]);
    }
    
    showScanConfirmation(modal, material) {
        // Create overlay confirmation
        const confirmation = document.createElement('div');
        confirmation.className = 'scan-confirmation';
        confirmation.innerHTML = `
            <div class="confirmation-content">
                <i class="fas fa-check-circle"></i>
                <h4>Material Scanned Successfully!</h4>
                <p>${material.name} (${material.category})</p>
                <div class="confirmation-actions">
                    <button class="btn btn-secondary" onclick="this.closest('.scan-confirmation').remove()">
                        Scan Another
                    </button>
                    <button class="btn btn-primary" onclick="scanner.useScannedMaterial(this.closest('.modal'))">
                        Use Now
                    </button>
                </div>
            </div>
        `;
        
        modal.querySelector('.scanner-container').appendChild(confirmation);
    }
    
    useScannedMaterial(modal) {
        const materialName = modal.querySelector('.material-name-display').textContent;
        const material = this.scanResults[this.scanResults.length - 1]?.materials[0];
        
        if (material) {
            // Add to material library
            if (window.materialLibrary) {
                window.materialLibrary.push(material);
                if (window.renderMaterialGrid) {
                    window.renderMaterialGrid(window.materialLibrary);
                }
            }
            
            // Select the material
            if (window.selectMaterial) {
                window.selectMaterial(material);
            }
            
            // Show notification
            if (window.showNotification) {
                window.showNotification(`Added to library: ${material.name}`, 'success');
            }
            
            console.log('✅ Material added to library:', material);
        }
        
        // Close modal
        modal.remove();
        this.stopScanning();
    }
    
    // ====== HELPER FUNCTIONS ======
    
    rgbToHex(r, g, b) {
        return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);
    }
    
    hexToRgb(hex) {
        const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
        return result ? {
            r: parseInt(result[1], 16),
            g: parseInt(result[2], 16),
            b: parseInt(result[3], 16)
        } : null;
    }
    
    // ====== PUBLIC API ======
    
    getScanResults() {
        return this.scanResults;
    }
    
    getLastScan() {
        return this.scanResults[this.scanResults.length - 1];
    }
    
    clearScanResults() {
        this.scanResults = [];
    }
    
    async scanImageFile(file) {
        return await this.processImageFile(file);
    }
    
    async scanFromCamera() {
        const modal = await this.openScannerModal();
        return modal;
    }
    
    exportScanData(format = 'json') {
        const data = this.getLastScan();
        
        switch (format) {
            case 'json':
                return JSON.stringify(data, null, 2);
            case 'csv':
                return this.convertToCSV(data);
            case 'xml':
                return this.convertToXML(data);
            default:
                return data;
        }
    }
    
    convertToCSV(data) {
        if (!data.materials || data.materials.length === 0) return '';
        
        const material = data.materials[0];
        const headers = ['Name', 'Category', 'Density', 'Friction', 'Elasticity', 'Confidence'];
        const values = [
            material.name,
            material.category,
            material.properties.density,
            material.properties.friction,
            material.properties.elasticity,
            material.aiConfidence
        ];
        
        return [headers.join(','), values.join(',')].join('\n');
    }
    
    convertToXML(data) {
        if (!data.materials || data.materials.length === 0) return '';
        
        const material = data.materials[0];
        return `
<?xml version="1.0" encoding="UTF-8"?>
<materialScan>
    <timestamp>${new Date(data.timestamp).toISOString()}</timestamp>
    <material>
        <name>${material.name}</name>
        <category>${material.category}</category>
        <confidence>${material.aiConfidence}</confidence>
        <properties>
            <density>${material.properties.density}</density>
            <friction>${material.properties.friction}</friction>
            <elasticity>${material.properties.elasticity}</elasticity>
            <color>${material.properties.color}</color>
        </properties>
        <tags>
            ${material.tags.map(tag => `<tag>${tag}</tag>`).join('')}
        </tags>
    </material>
</materialScan>
        `.trim();
    }
}

// ====== GLOBAL SCANNER INSTANCE ======

let scanner = null;

// Initialize scanner when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    scanner = new MaterialScannerAI();
    console.log('📸 Material Scanner ready');
    
    // Make scanner globally available for UI buttons
    window.scanner = scanner;
    
    // Add scanner button event listeners if they exist
    const scanButtons = document.querySelectorAll('[data-action="scan-material"]');
    scanButtons.forEach(btn => {
        btn.addEventListener('click', () => scanner.scanFromCamera());
    });
});

// ====== SCANNER UI STYLES ======

const scannerStyles = document.createElement('style');
scannerStyles.textContent = `
    .scanner-modal {
        max-width: 900px;
        padding: 0;
        overflow: hidden;
    }
    
    .scanner-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 20px;
        background: linear-gradient(135deg, var(--dark) 0%, var(--dark-light) 100%);
        color: white;
    }
    
    .scanner-header h3 {
        margin: 0;
        display: flex;
        align-items: center;
        gap: 10px;
    }
    
    .btn-close {
        background: rgba(255,255,255,0.1);
        border: none;
        color: white;
        width: 36px;
        height: 36px;
        border-radius: 50%;
        cursor: pointer;
        display: flex;
        align-items: center;
        justify-content: center;
        transition: all 0.2s ease;
    }
    
    .btn-close:hover {
        background: rgba(255,255,255,0.2);
        transform: rotate(90deg);
    }
    
    .scanner-container {
        max-height: 70vh;
        overflow-y: auto;
    }
    
    .scan-confirmation {
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: rgba(0,0,0,0.9);
        display: flex;
        align-items: center;
        justify-content: center;
        z-index: 100;
        animation: fadeIn 0.3s ease;
    }
    
    .confirmation-content {
        background: white;
        padding: 30px;
        border-radius: var(--radius-lg);
        text-align: center;
        max-width: 400px;
        animation: scaleIn 0.3s ease;
    }
    
    .confirmation-content i {
        font-size: 48px;
        color: var(--success);
        margin-bottom: 15px;
    }
    
    .confirmation-content h4 {
        margin: 0 0 10px 0;
        color: var(--dark);
    }
    
    .confirmation-content p {
        color: var(--gray);
        margin-bottom: 20px;
    }
    
    .confirmation-actions {
        display: flex;
        gap: 10px;
        justify-content: center;
    }
    
    @keyframes fadeIn {
        from { opacity: 0; }
        to { opacity: 1; }
    }
    
    @keyframes scaleIn {
        from {
            opacity: 0;
            transform: scale(0.9);
        }
        to {
            opacity: 1;
            transform: scale(1);
        }
    }
    
    .scanner-status {
        display: flex;
        align-items: center;
        gap: 8px;
    }
    
    .status-dot {
        width: 8px;
        height: 8px;
        border-radius: 50%;
        background: var(--gray-light);
    }
    
    .status-dot.active {
        background: var(--success);
        animation: pulse 1.5s infinite;
    }
    
    .confidence-display {
        display: flex;
        align-items: center;
        gap: 10px;
    }
    
    .confidence-bar {
        flex: 1;
        height: 8px;
        background: var(--gray-lighter);
        border-radius: 4px;
        overflow: hidden;
    }
    
    .confidence-fill {
        height: 100%;
        background: linear-gradient(90deg, var(--warning), var(--success));
        transition: width 0.5s ease;
    }
    
    .confidence-value {
        font-weight: 600;
        color: var(--dark);
        min-width: 40px;
    }
    
    .properties-display {
        display: flex;
        gap: 5px;
        flex-wrap: wrap;
    }
    
    .property-tag {
        background: var(--gray-lighter);
        padding: 4px 8px;
        border-radius: 4px;
        font-size: 12px;
        color: var(--gray);
    }
    
    .scanner-actions {
        display: flex;
        gap: 10px;
        justify-content: flex-end;
        padding-top: 20px;
        border-top: 1px solid var(--gray-light);
        margin-top: 20px;
    }
    
    .ai-tags-container {
        margin-top: 15px;
    }
    
    .ai-tags-container label {
        display: block;
        font-size: 14px;
        color: var(--gray);
        margin-bottom: 8px;
        font-weight: 500;
    }
    
    .ai-tags {
        display: flex;
        gap: 5px;
        flex-wrap: wrap;
    }
`;

document.head.appendChild(scannerStyles);

console.log('🤖 Material Scanner module loaded');