const fileInput = document.getElementById('fileInput');
const skcInput = document.getElementById('skcInput');
const dashboard = document.getElementById('dashboard');
const printBtn = document.getElementById('printBtn');
const imageBtn = document.getElementById('imageBtn');
const debugMsg = document.getElementById('debug-msg');
const skcTableBody = document.getElementById('skc-table-body');
const improvementTableBody = document.getElementById('improvement-table-body');

// Store global values for calculations
let globalGMV = 0;
let globalSKC = 0;
// Store all SKC data for lookup
let allSKCData = [];

// Function to parse 8-digit date string (YYYYMMDD) to formatted date
function parseDateString(dateStr) {
    if (!dateStr || dateStr.length !== 8) return null;
    const year = dateStr.substring(0, 4);
    const month = dateStr.substring(4, 6);
    const day = dateStr.substring(6, 8);
    return { year, month, day };
}

// Function to format date as month/day/year
function formatDateRange(year, month, day) {
    return `${month}/${day}/${year}`;
}

// Function to update the report date stamp
function updateReportDate(dataRows) {
    const dateElement = document.getElementById('report-date');
    if (!dateElement) return;
    
    // Get current date (without time)
    const now = new Date();
    const reportDateString = now.toLocaleDateString('en-US', { 
        year: 'numeric', 
        month: 'long', 
        day: 'numeric'
    });
    
    let dateRangeText = '';
    
    // If dataRows provided, extract date range from seller data
    if (dataRows && dataRows.length > 0) {
        const latest = dataRows[0];
        const previous = dataRows[1] || null;
        
        // Get date from latest row
        const latestDateStr = getFuzzy(latest, '商家分层日期 Date') || 
                             latest['商家分层日期 Date'] || 
                             latest['商家分层日期'] ||
                             '';
        
        if (latestDateStr) {
            const latestDate = parseDateString(String(latestDateStr));
            if (latestDate) {
                const latestFormatted = formatDateRange(latestDate.year, latestDate.month, latestDate.day);
                
                if (previous) {
                    // Get date from previous row
                    const prevDateStr = getFuzzy(previous, '商家分层日期 Date') || 
                                       previous['商家分层日期 Date'] || 
                                       previous['商家分层日期'] ||
                                       '';
                    
                    if (prevDateStr) {
                        const prevDate = parseDateString(String(prevDateStr));
                        if (prevDate) {
                            const prevFormatted = formatDateRange(prevDate.year, prevDate.month, prevDate.day);
                            dateRangeText = ` using data from ${prevFormatted} - ${latestFormatted}`;
                        } else {
                            dateRangeText = ` using data from ${latestFormatted}`;
                        }
                    } else {
                        dateRangeText = ` using data from ${latestFormatted}`;
                    }
                } else {
                    dateRangeText = ` using data from ${latestFormatted}`;
                }
            }
        }
    }
    
    const dateSpan = dateElement.querySelector('span');
    if (dateSpan) {
        dateSpan.textContent = reportDateString + dateRangeText;
    }
}

// Helper function to extract sales revenue from a row (defined globally)
window.getSalesRevenue = function(row) {
    if (!row) return 0;
    
    const keys = Object.keys(row);
    let salesRevenueValue = null;
    
    // First, try case-insensitive exact match (most reliable)
    const targetKey = 'sales revenue within 2-31 days';
    const exactKey = keys.find(k => k.toLowerCase().trim() === targetKey);
    if (exactKey) {
        salesRevenueValue = row[exactKey];
    }
    
    // If still not found, try direct property access (case-sensitive)
    if ((salesRevenueValue === null || salesRevenueValue === undefined || salesRevenueValue === '') && row.hasOwnProperty('Sales revenue within 2-31 days')) {
        salesRevenueValue = row['Sales revenue within 2-31 days'];
    }
    
    // Try with trailing space
    if ((salesRevenueValue === null || salesRevenueValue === undefined || salesRevenueValue === '') && row.hasOwnProperty('Sales revenue within 2-31 days ')) {
        salesRevenueValue = row['Sales revenue within 2-31 days '];
    }
    
    // If still not found, search through all keys (prioritize exact match pattern)
    if (salesRevenueValue === null || salesRevenueValue === undefined || salesRevenueValue === '') {
        // Look for key containing 'sales revenue' and '2-31' and 'within', but NOT 'total'
        let revenueKey = keys.find(k => {
            const lowerK = k.toLowerCase().trim();
            return lowerK.includes('sales') && 
                   lowerK.includes('revenue') && 
                   (lowerK.includes('2-31') || lowerK.includes('within')) && 
                   !lowerK.includes('total');
        });
        
        // If still not found, try just 'revenue' and '2-31' or 'within', but not 'total'
        if (!revenueKey) {
            revenueKey = keys.find(k => {
                const lowerK = k.toLowerCase().trim();
                return lowerK.includes('revenue') && 
                       (lowerK.includes('2-31') || lowerK.includes('within')) && 
                       !lowerK.includes('total');
            });
        }
        
        salesRevenueValue = revenueKey ? row[revenueKey] : null;
    }
    
    // If still not found, use fuzzy search as last resort
    if (salesRevenueValue === null || salesRevenueValue === undefined || salesRevenueValue === '') {
        if (typeof getFuzzy !== 'undefined') {
            salesRevenueValue = getFuzzy(row, 'Sales revenue within 2-31 days');
        }
    }
    
    // Parse the value - handle both string and number formats
    let salesRevenue = 0;
    if (salesRevenueValue !== null && salesRevenueValue !== undefined && salesRevenueValue !== '') {
        const strValue = String(salesRevenueValue).replace(/,/g, '').trim();
        const parsed = parseFloat(strValue);
        salesRevenue = isNaN(parsed) ? 0 : parsed;
    }
    return salesRevenue;
};

// Tooltip positioning function
function positionTooltip(wrapper, tooltip) {
    const rect = wrapper.getBoundingClientRect();
    const tooltipWidth = 300;
    const tooltipHeight = tooltip.offsetHeight || 150;
    const spacing = 8;
    
    let top = rect.top - tooltipHeight - spacing;
    let left = rect.left + (rect.width / 2) - (tooltipWidth / 2);
    
    // Adjust if tooltip goes off top of screen
    if (top < 10) {
        top = rect.bottom + spacing;
        tooltip.classList.add('tooltip-below');
    } else {
        tooltip.classList.remove('tooltip-below');
    }
    
    // Adjust if tooltip goes off left edge
    if (left < 10) {
        left = 10;
    }
    
    // Adjust if tooltip goes off right edge
    if (left + tooltipWidth > window.innerWidth - 10) {
        left = window.innerWidth - tooltipWidth - 10;
    }
    
    tooltip.style.top = top + 'px';
    tooltip.style.left = left + 'px';
}

// Track active tooltip
let activeTooltip = null;
let activeWrapper = null;

// Initialize tooltip positioning
function initTooltips() {
    const tooltipWrappers = document.querySelectorAll('.tooltip-wrapper');
    tooltipWrappers.forEach(wrapper => {
        const tooltip = wrapper.querySelector('.tooltip-content');
        if (tooltip && !wrapper.dataset.tooltipInitialized) {
            wrapper.dataset.tooltipInitialized = 'true';
            wrapper.addEventListener('mouseenter', () => {
                activeTooltip = tooltip;
                activeWrapper = wrapper;
                // Force a reflow to get accurate height
                tooltip.style.visibility = 'visible';
                tooltip.style.opacity = '0';
                // Use requestAnimationFrame to ensure layout is calculated
                requestAnimationFrame(() => {
                    positionTooltip(wrapper, tooltip);
                    tooltip.style.opacity = '1';
                });
            });
            wrapper.addEventListener('mouseleave', () => {
                tooltip.style.visibility = 'hidden';
                tooltip.style.opacity = '0';
                activeTooltip = null;
                activeWrapper = null;
            });
        }
    });
}

// Initialize on DOM ready
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', initTooltips);
} else {
    initTooltips();
}

// Re-initialize tooltips when SKC data is loaded (for dynamically added content)
const originalProcessSKCData = window.processSKCData;
if (typeof originalProcessSKCData === 'undefined') {
    // Will be set later, but we'll also call initTooltips after table updates
}

// Update tooltip positions on scroll and resize
window.addEventListener('scroll', () => {
    if (activeTooltip && activeWrapper) {
        positionTooltip(activeWrapper, activeTooltip);
    }
}, true);

window.addEventListener('resize', () => {
    if (activeTooltip && activeWrapper) {
        positionTooltip(activeWrapper, activeTooltip);
    }
});

printBtn.addEventListener('click', () => {
    window.focus();
    window.print();
});

// Image Export Function
imageBtn.addEventListener('click', async () => {
    // Show loading state
    const originalText = imageBtn.textContent;
    imageBtn.textContent = 'Generating Image...';
    imageBtn.disabled = true;
    
    try {
        // Hide elements that shouldn't be in the image (no-print class)
        const noPrintElements = document.querySelectorAll('.no-print');
        noPrintElements.forEach(el => {
            el.style.display = 'none';
        });
        
        // Use html2canvas to capture the dashboard
        const canvas = await html2canvas(dashboard, {
            backgroundColor: '#ffffff',
            scale: 2, // Higher quality
            logging: false,
            useCORS: true,
            allowTaint: true,
            windowWidth: dashboard.scrollWidth,
            windowHeight: dashboard.scrollHeight
        });
        
        // Restore no-print elements
        noPrintElements.forEach(el => {
            el.style.display = '';
        });
        
        // Convert canvas to blob and download
        canvas.toBlob((blob) => {
            const url = URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.href = url;
            
            // Generate filename with current date
            const now = new Date();
            const dateStr = now.toISOString().split('T')[0];
            const shopName = document.getElementById('ui-shop-name')?.textContent?.replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_]/g, '') || 'report';
            link.download = `SHEIN_AM_Dashboard_${shopName}_${dateStr}.png`;
            
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(url);
        }, 'image/png');
        
    } catch (error) {
        console.error('Error generating image:', error);
        alert('Error generating image. Please try again.');
    } finally {
        // Restore button state
        imageBtn.textContent = originalText;
        imageBtn.disabled = false;
    }
});

// Function to find SKC data by SKC ID
function findSKCData(skcId) {
    if (!skcId || !allSKCData || allSKCData.length === 0) return null;
    
    const skcIdStr = String(skcId).trim();
    return allSKCData.find(row => {
        const rowSKC = getFuzzy(row, 'SKC');
        return rowSKC && String(rowSKC).trim() === skcIdStr;
    }) || null;
}

// Function to populate improvement row data from SKC
function populateImprovementRow(skcId, rowElement) {
    const skcData = findSKCData(skcId);
    if (!skcData) {
        // Clear fields if SKC not found
        rowElement.querySelector('.improvement-exposure').textContent = '--';
        rowElement.querySelector('.improvement-ctr').textContent = '--';
        rowElement.querySelector('.improvement-cvr').textContent = '--';
        rowElement.querySelector('.improvement-sales').textContent = '--';
        return;
    }
    
    // Get Exposure (UV) - now 2-31 days
    const uvValue = getFuzzy(skcData, '2~31 days UV exposure') || 
                   getFuzzy(skcData, '2~15 days of UV exposure') || 
                   getFuzzy(skcData, 'UV');
    const uv = uvValue !== null && uvValue !== undefined && uvValue !== '' ? parseInt(uvValue) : null;
    rowElement.querySelector('.improvement-exposure').textContent = (uv !== null && !isNaN(uv)) ? uv.toLocaleString() : '--';
    
    // Get CTR - still 15 days
    const ctrValue = skcData['CTR in the past 15 days'] || 
                    skcData['CTR in the past 15 days '] ||
                    getFuzzy(skcData, 'CTR in the past 15 days') || 
                    getFuzzy(skcData, 'CTR');
    const ctr = (ctrValue !== null && ctrValue !== undefined && ctrValue !== '') ? parsePercentage(ctrValue) : null;
    rowElement.querySelector('.improvement-ctr').textContent = (ctr !== null && !isNaN(ctr)) ? ctr.toFixed(2) + '%' : '--';
    
    // Get CVR - now 30 days
    let cvrValue = skcData['Convertion rate from Product Detail Page in the past 30 days'] ||
                   skcData['Convertion rate from Product Detail Page in the past 30 days '] ||
                   skcData['Convertion rate from Product Detail Page in the past 15 days'] ||
                   skcData['Convertion rate from Product Detail Page in the past 15 days '] ||
                   getFuzzy(skcData, 'Convertion rate from Product Detail Page in the past 30 days') ||
                   getFuzzy(skcData, 'Convertion rate from Product Detail Page in the past 15 days') || 
                   getFuzzy(skcData, 'Convertion rate') || 
                   getFuzzy(skcData, 'Conversion rate') ||
                   getFuzzy(skcData, 'CVR');
    
    if (cvrValue === null || cvrValue === undefined || cvrValue === '') {
        const keys = Object.keys(skcData || {});
        // Prioritize 30 days over 15 days
        const cvrKey = keys.find(k => {
            const lowerK = k.toLowerCase();
            return (lowerK.includes('convertion') || lowerK.includes('conversion')) && 
                   lowerK.includes('rate') && 
                   lowerK.includes('30') &&
                   !lowerK.includes('return');
        }) || keys.find(k => {
            const lowerK = k.toLowerCase();
            return (lowerK.includes('convertion') || lowerK.includes('conversion')) && 
                   lowerK.includes('rate') && 
                   !lowerK.includes('return');
        }) || keys.find(k => {
            const lowerK = k.toLowerCase();
            return lowerK.includes('cvr') && !lowerK.includes('return');
        });
        cvrValue = cvrKey ? skcData[cvrKey] : null;
    }
    
    const cvr = (cvrValue !== null && cvrValue !== undefined && cvrValue !== '') ? parsePercentage(cvrValue) : null;
    rowElement.querySelector('.improvement-cvr').textContent = (cvr !== null && !isNaN(cvr)) ? cvr.toFixed(2) + '%' : '--';
    
    // Get Sales Revenue
    const salesRevenue = getSalesRevenue(skcData);
    const salesRevenueFormatted = (salesRevenue !== null && !isNaN(salesRevenue)) ? '$' + salesRevenue.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '--';
    rowElement.querySelector('.improvement-sales').textContent = salesRevenueFormatted;
}

// Function to add a new row to improvement table
function addImprovementRow() {
    const row = document.createElement('tr');
    row.className = "border-b hover:bg-gray-50 transition";
    row.innerHTML = `
        <td class="p-2 border">
            <input type="text" class="improvement-skc-input w-full text-xs border-none bg-transparent focus:outline-none focus:ring-1 focus:ring-blue-500 px-1" placeholder="Enter SKC...">
        </td>
        <td class="p-2 border text-center improvement-ctr">--</td>
        <td class="p-2 border text-center improvement-cvr">--</td>
        <td class="p-2 border text-center improvement-exposure">--</td>
        <td class="p-2 border text-center improvement-sales">--</td>
        <td class="p-2 border notes-cell">
            <textarea class="w-full text-xs text-gray-500 border-none bg-transparent focus:outline-none focus:ring-0" placeholder="Add notes..." rows="1"></textarea>
        </td>
        <td class="p-2 border">
            <button onclick="this.closest('tr').remove()" class="no-print bg-red-500 text-white text-[10px] font-bold px-2 py-1 rounded hover:bg-red-700 transition">Delete</button>
        </td>
    `;
    
    improvementTableBody.appendChild(row);
    
    // Add event listener to SKC input
    const skcInput = row.querySelector('.improvement-skc-input');
    skcInput.addEventListener('blur', () => {
        const skcId = skcInput.value.trim();
        if (skcId) {
            populateImprovementRow(skcId, row);
        }
    });
    
    skcInput.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            skcInput.blur();
        }
    });
}

// Generic Excel Reader
function readExcel(file, callback) {
    const reader = new FileReader();
    reader.onload = (evt) => {
        try {
            const data = evt.target.result;
            const workbook = XLSX.read(data, { type: 'binary' });
            const firstSheetName = workbook.SheetNames[0];
            const rawData = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName]);
            callback(rawData);
        } catch (err) {
            showError("Error processing file.");
        }
    };
    reader.readAsBinaryString(file);
}

// 1. Performance File
fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;
    readExcel(file, (rawData) => {
        if (rawData && rawData.length >= 1) {
            debugMsg.classList.add('hidden');
            processSellerData(rawData);
        } else {
            showError("Requires at least Header and 1 data row.");
        }
    });
});

// 2. SKC Data File
skcInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;
    readExcel(file, (rawData) => {
        if (rawData && rawData.length >= 1) {
            debugMsg.classList.add('hidden');
            processSKCData(rawData);
        } else {
            showError("SKC file appears empty.");
        }
    });
});

function showError(msg) {
    debugMsg.textContent = msg;
    debugMsg.classList.remove('hidden');
}

function getFuzzy(obj, searchKey) {
    if (!obj) return null;
    const keys = Object.keys(obj);
    // Try exact match first (case-insensitive)
    const exactKey = keys.find(k => k.toLowerCase() === searchKey.toLowerCase());
    if (exactKey) return obj[exactKey];
    // Then try substring match
    const foundKey = keys.find(k => k.toLowerCase().includes(searchKey.toLowerCase()));
    return foundKey ? obj[foundKey] : null;
}

function parsePercentage(value) {
    // Handle null, undefined, empty string
    if (value === null || value === undefined || value === '') return 0;
    
    // Convert to string for processing
    const strValue = String(value).trim();
    if (strValue === '' || strValue === 'NaN') return 0;
    
    // If it's already a string with %, parse it as percentage
    if (strValue.includes('%')) {
        const num = parseFloat(strValue.replace(/%/g, '').trim());
        return isNaN(num) ? 0 : num;
    }
    
    // Parse as number
    const num = parseFloat(strValue);
    if (isNaN(num)) return 0;
    
    // If it's a number less than 1 and greater than 0, assume it's a decimal (0.0069 = 0.69%)
    if (num < 1 && num > 0) {
        return num * 100;
    }
    
    // Otherwise assume it's already a percentage value (e.g., 0.69 means 0.69%)
    return num;
}

function processSellerData(dataRows) {
    const latest = dataRows[0];
    const previous = dataRows[1] || null;

    dashboard.style.opacity = "1";
    dashboard.style.pointerEvents = "all";
    updateReportDate(dataRows);

    document.getElementById('ui-am-cat').textContent = getFuzzy(latest, '商家运营') || 'Unassigned';
    const shopName = getFuzzy(latest, '店铺名称') || 'N/A';
    const shopId = latest.supplier_id || 'N/A';
    document.getElementById('ui-shop-name').textContent = `${shopName} (#${shopId})`;
    document.getElementById('ui-category').textContent = getFuzzy(latest, '供应商一级分类') || 'N/A';
    document.getElementById('ui-top-cat').textContent = getFuzzy(latest, 'Top1品类') || 'N/A';
    document.getElementById('ui-top2-category').textContent = (getFuzzy(latest, 'Top2品类') || 'N/A').toLowerCase();
    document.getElementById('ui-tier').textContent = getFuzzy(latest, '月度商家分层') || 'TBD';

    const currGMV = parseFloat(getFuzzy(latest, '近30d销额') || 0);
    const curr90dGMV = parseFloat(getFuzzy(latest, '近90d销额') || 0);
    const currSKC = parseInt(getFuzzy(latest, '在售动销数') || 0);
    
    // Store globally for use in SKC calculations
    globalGMV = currGMV;
    globalSKC = currSKC;
    
    // Calculate and display Average Revenue per SKC
    const avgRevPerSKC = currSKC > 0 ? currGMV / currSKC : 0;
    document.getElementById('ui-avg-rev-per-skc').textContent = `$${avgRevPerSKC.toFixed(2)}`;
    
    document.getElementById('ui-30d-gmv').textContent = `$${currGMV.toLocaleString()}`;
    document.getElementById('ui-90d-gmv').textContent = `$${curr90dGMV.toLocaleString()}`;

    const prevGMV = previous ? parseFloat(getFuzzy(previous, '近30d销额') || 0) : 0;
    const prevSKC = previous ? parseInt(getFuzzy(previous, '在售动销数') || 0) : 0;

    document.getElementById('comp-curr-gmv').textContent = `$${currGMV.toLocaleString()}`;
    document.getElementById('comp-prev-gmv').textContent = previous ? `$${prevGMV.toLocaleString()}` : 'N/A';
    document.getElementById('comp-curr-skc').textContent = currSKC;
    document.getElementById('comp-prev-skc').textContent = previous ? prevSKC : 'N/A';

    if (previous) {
        const gmvGrowth = prevGMV > 0 ? ((currGMV - prevGMV) / prevGMV) * 100 : 0;
        const skcGrowth = currSKC - prevSKC;
        updateTrend('comp-trend-gmv', gmvGrowth, '%');
        updateTrend('comp-trend-skc', skcGrowth, ' units');
    }

    const listingsCount = parseInt(getFuzzy(latest, 't-30~t-1上新数') || 0);
    updateKPI('data-new-listings', 'pill-listings', listingsCount, 10, 'more', ' SKCs');
    updateKPI('data-active-skc', 'pill-active-skc', currSKC, 30, 'more', ' SKCs');

    const avail = parseInt(getFuzzy(latest, 't-30~t-1活动可报名数') || 0);
    const done = parseInt(getFuzzy(latest, 't-30~t-1活动已报名数') || 0);
    const eventRate = avail > 0 ? (done / avail) * 100 : 0;
    updateKPI('data-events', 'pill-events', eventRate, 50, 'more', '%');

    const den72 = parseFloat(getFuzzy(latest, '72h揽收率分母') || 0);
    const mol72 = parseFloat(getFuzzy(latest, '72h揽收率分子') || 0);
    const rate72 = den72 > 0 ? (mol72 / den72) * 100 : 0;
    updateKPI('data-72h', 'pill-72h', rate72, 90, 'more', '%');

    const denLS = parseFloat(getFuzzy(latest, '断码销额占比分母') || 0);
    const molLS = parseFloat(getFuzzy(latest, '断码销额占比分子') || 0);
    const rateLS = denLS > 0 ? (molLS / denLS) * 100 : 0;
    updateKPI('data-lost-sales', 'pill-lost-sales', rateLS, 8, 'less', '%');
}

function processSKCData(dataRows) {
    // Store all SKC data globally for lookup in improvement table
    allSKCData = dataRows;
    
    // Collect all product tiers with their metrics
    const tierData = new Map(); // Map<tier, {count, uvSum, ctrSum, cvrSum, uvCount, ctrCount, cvrCount}>
    
    dataRows.forEach(row => {
        const level = getFuzzy(row, 'Product level (current)') || '-';
        if (level && level !== '-' && level.trim() !== '') {
            // Initialize tier data if not exists
            if (!tierData.has(level)) {
                tierData.set(level, {
                    count: 0,
                    uvSum: 0,
                    ctrSum: 0,
                    cvrSum: 0,
                    uvCount: 0,
                    ctrCount: 0,
                    cvrCount: 0
                });
            }
            
            const tierInfo = tierData.get(level);
            tierInfo.count++;
            
            // Get UV Exposure (now 2~31 days UV exposure)
            const uv = parseInt(getFuzzy(row, '2~31 days UV exposure') || 
                               getFuzzy(row, '2~15 days of UV exposure') || 
                               getFuzzy(row, 'UV') || 0);
            if (uv && !isNaN(uv)) {
                tierInfo.uvSum += uv;
                tierInfo.uvCount++;
            }
            
            // Get CTR (CTR in the past 15 days - still 15 days)
            const ctrValue = row['CTR in the past 15 days'] || 
                            row['CTR in the past 15 days '] ||
                            getFuzzy(row, 'CTR in the past 15 days') || 
                            getFuzzy(row, 'CTR') || 
                            0;
            const ctr = parsePercentage(ctrValue);
            if (ctr && !isNaN(ctr) && ctr > 0) {
                tierInfo.ctrSum += ctr;
                tierInfo.ctrCount++;
            }
            
            // Get CVR (now Convertion rate from Product Detail Page in the past 30 days)
            let cvrValue = row['Convertion rate from Product Detail Page in the past 30 days'] ||
                           row['Convertion rate from Product Detail Page in the past 30 days '] ||
                           row['Convertion rate from Product Detail Page in the past 15 days'] ||
                           row['Convertion rate from Product Detail Page in the past 15 days '] ||
                           getFuzzy(row, 'Convertion rate from Product Detail Page in the past 30 days') ||
                           getFuzzy(row, 'Convertion rate from Product Detail Page in the past 15 days') || 
                           getFuzzy(row, 'Convertion rate') || 
                           getFuzzy(row, 'Conversion rate') ||
                           getFuzzy(row, 'CVR');
            
            // If still not found, search through all keys (prioritize 30 days)
            if (cvrValue === null || cvrValue === undefined || cvrValue === '') {
                const keys = Object.keys(row || {});
                const cvrKey = keys.find(k => {
                    const lowerK = k.toLowerCase();
                    return (lowerK.includes('convertion') || lowerK.includes('conversion')) && 
                           lowerK.includes('rate') && 
                           lowerK.includes('30') &&
                           !lowerK.includes('return');
                }) || keys.find(k => {
                    const lowerK = k.toLowerCase();
                    return (lowerK.includes('convertion') || lowerK.includes('conversion')) && 
                           lowerK.includes('rate') && 
                           !lowerK.includes('return');
                }) || keys.find(k => {
                    const lowerK = k.toLowerCase();
                    return lowerK.includes('cvr') && !lowerK.includes('return');
                });
                cvrValue = cvrKey ? row[cvrKey] : 0;
            }
            
            const cvr = parsePercentage(cvrValue);
            if (cvr && !isNaN(cvr) && cvr > 0) {
                tierInfo.cvrSum += cvr;
                tierInfo.cvrCount++;
            }
        }
    });
    
    // Populate the tier summary table
    const tierSummaryBody = document.getElementById('tier-summary-table-body');
    tierSummaryBody.innerHTML = '';
    
    // Convert Map to Array and sort by number of products per tier (descending order)
    const sortedTiers = Array.from(tierData.keys()).sort((a, b) => {
        const countA = tierData.get(a).count;
        const countB = tierData.get(b).count;
        return countB - countA; // Descending order
    });
    
    // If no tiers found, show a message
    if (sortedTiers.length === 0) {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td colspan="6" class="p-8 text-center text-gray-400 italic">No product tier data available.</td>
        `;
        tierSummaryBody.appendChild(tr);
    } else {
        // Create rows for each unique tier dynamically
        sortedTiers.forEach(tier => {
            const tierInfo = tierData.get(tier);
            const count = tierInfo.count;
            
            // Calculate averages
            const avgUV = tierInfo.uvCount > 0 ? (tierInfo.uvSum / tierInfo.uvCount) : 0;
            const avgCTR = tierInfo.ctrCount > 0 ? (tierInfo.ctrSum / tierInfo.ctrCount) : 0;
            const avgCVR = tierInfo.cvrCount > 0 ? (tierInfo.cvrSum / tierInfo.cvrCount) : 0;
            
            const tr = document.createElement('tr');
            tr.className = "border-b hover:bg-gray-50 transition";
            tr.innerHTML = `
                <td class="p-2 border font-semibold">${tier}</td>
                <td class="p-2 border">${count}</td>
                <td class="p-2 border">${avgUV > 0 ? avgUV.toLocaleString() : '--'}</td>
                <td class="p-2 border">${avgCTR > 0 ? avgCTR.toFixed(2) + '%' : '--'}</td>
                <td class="p-2 border">${avgCVR > 0 ? avgCVR.toFixed(2) + '%' : '--'}</td>
                <td class="p-2 border notes-cell">
                    <textarea class="w-full text-xs text-gray-500 border-none bg-transparent focus:outline-none focus:ring-0" placeholder="Add notes..." rows="1"></textarea>
                </td>
            `;
            tierSummaryBody.appendChild(tr);
        });
    }
    
    
    // Sort dataRows by sales revenue (descending) before taking top 10
    const sortedDataRows = [...dataRows].sort((a, b) => {
        const revenueA = getSalesRevenue(a);
        const revenueB = getSalesRevenue(b);
        return revenueB - revenueA; // Descending order
    });
    
    // Keep top 10 for the detailed table
    const top10 = sortedDataRows.slice(0, 10);
    skcTableBody.innerHTML = '';

    top10.forEach(row => {
        const tr = document.createElement('tr');
        tr.className = "border-b hover:bg-gray-50 transition";
        
        const imgUrl = getFuzzy(row, 'picture') || '';
        const skcId = getFuzzy(row, 'SKC') || 'N/A';
        
        // Try multiple variations for product name
        let name = row['Product English name'] || 
                   row['Product English Name'] || 
                   row['product english name'] ||
                   getFuzzy(row, 'Product English name') || 
                   getFuzzy(row, 'English name') ||
                   getFuzzy(row, 'name') ||
                   getFuzzy(row, 'product') ||
                   (() => {
                       // Last resort: find any key containing 'name' or 'product'
                       const keys = Object.keys(row || {});
                       const nameKey = keys.find(k => 
                           k.toLowerCase().includes('english') && k.toLowerCase().includes('name')
                       ) || keys.find(k => k.toLowerCase().includes('name'));
                       return nameKey ? row[nameKey] : 'Unnamed Product';
                   })();
        
        const level = getFuzzy(row, 'Product level (current)') || '-';
        
        // Try multiple ways to get CTR (still 15 days)
        const ctrValue = row['CTR in the past 15 days'] || 
                        row['CTR in the past 15 days '] ||
                        getFuzzy(row, 'CTR in the past 15 days') || 
                        getFuzzy(row, 'CTR') || 
                        0;
        const ctr = parsePercentage(ctrValue).toFixed(2) + '%';
        
        // Try multiple ways to get CVR (now 30 days)
        let cvrValue = row['Convertion rate from Product Detail Page in the past 30 days'] ||
                       row['Convertion rate from Product Detail Page in the past 30 days '] ||
                       row['Convertion rate from Product Detail Page in the past 15 days'] ||
                       row['Convertion rate from Product Detail Page in the past 15 days '] ||
                       getFuzzy(row, 'Convertion rate from Product Detail Page in the past 30 days') ||
                       getFuzzy(row, 'Convertion rate from Product Detail Page in the past 15 days') || 
                       getFuzzy(row, 'Convertion rate') || 
                       getFuzzy(row, 'Conversion rate') ||
                       getFuzzy(row, 'CVR');
        
        // If still not found (null, undefined, or empty string), search through all keys
        if (cvrValue === null || cvrValue === undefined || cvrValue === '') {
            const keys = Object.keys(row || {});
            // Look for key containing 'convertion' or 'conversion' and 'rate', but not 'return'
            // Prioritize 30 days over 15 days
            const cvrKey = keys.find(k => {
                const lowerK = k.toLowerCase();
                return (lowerK.includes('convertion') || lowerK.includes('conversion')) && 
                       lowerK.includes('rate') && 
                       lowerK.includes('30') &&
                       !lowerK.includes('return');
            }) || keys.find(k => {
                const lowerK = k.toLowerCase();
                return (lowerK.includes('convertion') || lowerK.includes('conversion')) && 
                       lowerK.includes('rate') && 
                       !lowerK.includes('return');
            }) || keys.find(k => {
                const lowerK = k.toLowerCase();
                return lowerK.includes('cvr') && !lowerK.includes('return');
            });
            cvrValue = cvrKey ? row[cvrKey] : 0;
        }
        
        const cvr = parsePercentage(cvrValue).toFixed(2) + '%';
        
        // Get UV (now 2~31 days UV exposure)
        const uv = parseInt(getFuzzy(row, '2~31 days UV exposure') || 
                           getFuzzy(row, '2~15 days of UV exposure') || 
                           getFuzzy(row, 'UV') || 0).toLocaleString();
        
        // Get Sales Quantity within 2-31 days
        let salesQuantityValue = row['Sales within 2-31 days'] ||
                                row['Sales within 2-31 days '] ||
                                row['Quantity within 2-31 days'] ||
                                row['Quantity within 2-31 days '] ||
                                getFuzzy(row, 'Sales within 2-31 days') ||
                                getFuzzy(row, 'Quantity within 2-31 days') ||
                                getFuzzy(row, 'Sales within') ||
                                getFuzzy(row, 'Quantity within') ||
                                getFuzzy(row, 'Sales quantity') ||
                                getFuzzy(row, 'Quantity');
        
        // If still not found, search through all keys
        if (salesQuantityValue === null || salesQuantityValue === undefined || salesQuantityValue === '') {
            const keys = Object.keys(row || {});
            // Look for key containing 'sales' or 'quantity' and '2-31' or 'within', but not 'revenue' or 'total'
            let quantityKey = keys.find(k => {
                const lowerK = k.toLowerCase();
                return (lowerK.includes('sales') || lowerK.includes('quantity')) && 
                       (lowerK.includes('2-31') || lowerK.includes('within')) && 
                       !lowerK.includes('revenue') && 
                       !lowerK.includes('total');
            });
            // If still not found, try just 'quantity' and '2-31' or 'sales' and 'quantity' and '2-31'
            if (!quantityKey) {
                quantityKey = keys.find(k => {
                    const lowerK = k.toLowerCase();
                    return (lowerK.includes('quantity') && lowerK.includes('2-31')) ||
                           (lowerK.includes('sales') && lowerK.includes('quantity') && lowerK.includes('2-31'));
                });
            }
            // Last resort: find any column with 'quantity' but not 'return' or 'rate'
            if (!quantityKey) {
                quantityKey = keys.find(k => {
                    const lowerK = k.toLowerCase();
                    return lowerK.includes('quantity') && 
                           !lowerK.includes('return') && 
                           !lowerK.includes('rate') &&
                           !lowerK.includes('revenue');
                });
            }
            salesQuantityValue = quantityKey ? row[quantityKey] : 0;
        }
        
        // Parse the quantity value
        let salesQuantity = 0;
        if (salesQuantityValue !== null && salesQuantityValue !== undefined && salesQuantityValue !== '') {
            const parsed = parseInt(String(salesQuantityValue).replace(/,/g, ''));
            salesQuantity = isNaN(parsed) ? 0 : parsed;
        }
        const salesQuantityFormatted = salesQuantity.toLocaleString();
        
        // Get Sales Revenue within 2-31 days (use the shared function)
        const salesRevenue = getSalesRevenue(row);
        const salesRevenueFormatted = salesRevenue > 0 ? '$' + salesRevenue.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '$0.00';

        tr.innerHTML = `
            <td class="p-3 border">
                <img src="${imgUrl}" class="skc-img" alt="SKC" onerror="this.src='https://via.placeholder.com/40x52?text=?'">
            </td>
            <td class="p-3 border">
                <div class="font-semibold text-gray-900">${skcId}</div>
                <div class="text-xs text-gray-500 truncate w-32 mt-1">${name}</div>
            </td>
            <td class="p-3 border text-center font-semibold text-blue-600">${level}</td>
            <td class="p-3 border text-center">${ctr}</td>
            <td class="p-3 border text-center">${cvr}</td>
            <td class="p-3 border text-center font-medium">${uv}</td>
            <td class="p-3 border text-center font-semibold">${salesQuantityFormatted}</td>
            <td class="p-3 border text-center font-semibold text-green-600">${salesRevenueFormatted}</td>
        `;
        skcTableBody.appendChild(tr);
    });

    // Re-initialize tooltips for dynamically added content
    initTooltips();

    // Ensure dashboard is visible if only SKC data is uploaded first
    dashboard.style.opacity = "1";
    dashboard.style.pointerEvents = "all";
    // Update date without data range if seller data hasn't been loaded yet
    const dateElement = document.getElementById('report-date');
    if (dateElement && !dateElement.querySelector('span').textContent.includes('using data from')) {
        updateReportDate(null);
    }
}

function updateTrend(id, value, suffix) {
    const elem = document.getElementById(id);
    const sign = value > 0 ? '▲' : (value < 0 ? '▼' : '');
    elem.textContent = `${sign} ${Math.abs(value).toFixed(1)}${suffix}`;
    elem.className = `p-3 font-bold ${value >= 0 ? 'trend-up' : 'trend-down'}`;
}

function updateKPI(textId, pillId, value, benchmark, direction, unit) {
    const textElem = document.getElementById(textId);
    const pillElem = document.getElementById(pillId);
    textElem.textContent = `${value.toFixed(1)}${unit}`;
    let status = 'good';
    if (direction === 'more') {
        if (value < (benchmark * 0.7)) status = 'critical';
        else if (value < benchmark) status = 'fair';
    } else {
        if (value > (benchmark * 1.5)) status = 'critical';
        else if (value > benchmark) status = 'fair';
    }
    pillElem.textContent = status;
    pillElem.className = `status-pill status-${status}`;
}

function updateQualityEvalStatus() {
    const dataCell = document.getElementById('data-quality-eval');
    const pillElem = document.getElementById('pill-quality-eval');
    const value = parseFloat(dataCell.textContent.trim());
    
    if (isNaN(value) || dataCell.textContent.trim() === '--' || dataCell.textContent.trim() === '') {
        pillElem.textContent = '--';
        pillElem.className = 'status-pill';
        return;
    }
    
    let statusText = '';
    let statusClass = '';
    
    if (value < 4) {
        statusText = 'Poor';
        statusClass = 'status-pill status-critical';
    } else if (value < 4.7) {
        statusText = 'Not met';
        statusClass = 'status-pill status-fair';
    } else {
        statusText = 'Good';
        statusClass = 'status-pill status-good';
    }
    
    pillElem.textContent = statusText;
    pillElem.className = statusClass;
}

// Initialize quality evaluation event listener
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
        const qualityEvalCell = document.getElementById('data-quality-eval');
        if (qualityEvalCell) {
            qualityEvalCell.addEventListener('blur', updateQualityEvalStatus);
            qualityEvalCell.addEventListener('keypress', (e) => {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    qualityEvalCell.blur();
                }
            });
        }
    });
} else {
    const qualityEvalCell = document.getElementById('data-quality-eval');
    if (qualityEvalCell) {
        qualityEvalCell.addEventListener('blur', updateQualityEvalStatus);
        qualityEvalCell.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                qualityEvalCell.blur();
            }
        });
    }
}
