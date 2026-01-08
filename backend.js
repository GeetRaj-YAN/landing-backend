require('dotenv').config();
const http = require('http');
const fs = require('fs');
const path = require('path');
const axios = require('axios');

const CSV_FILE = path.join(__dirname, 'bookings.csv');
const TOKENS_FILE = path.join(__dirname, 'tokens.json');
const PORT = 3001;

// Zoho Config
const ZOHO_REGION = process.env.ZOHO_REGION || 'in';
const ZOHO_CLIENT_ID = process.env.ZOHO_CLIENT_ID;
const ZOHO_CLIENT_SECRET = process.env.ZOHO_CLIENT_SECRET;
const ZOHO_GRANT_CODE = process.env.ZOHO_GRANT_CODE;
const ZOHO_REFRESH_TOKEN = process.env.ZOHO_REFRESH_TOKEN;
const ZOHO_WORKBOOK_ID = process.env.ZOHO_WORKBOOK_ID;
const ZOHO_SHEET_NAME = process.env.ZOHO_SHEET_NAME || 'Sheet1';

const FIELD_MAPPINGS = {
    'timestamp': ['Date', 'timestamp', 'date'],
    'quotationNo': ['Quotation No', 'quotation_no'],
    'name': ['Full Name', 'Name', 'name'],
    'phone': ['Contact No', 'Phone', 'phone'],
    'email': ['Mail Id', 'Email', 'email'],
    'city': ['Location', 'City', 'city'],
    'address': ['Full Address', 'address'],
    'mainService': ['Service', 'service'],
    'onlineService': ['Online Service'],
    'budget': ['Budget'],
    'plotLength': ['Plot Length'],
    'plotWidth': ['Plot Width'],
    'archPlotArea': ['Plot Area'],
    'archSetback': ['Setback'],
    'typicalAreaAfterSetback': ['Typical Area'],
    'archFloors': ['Floors'],
    'archTerrace': ['Terrace'],
    'archTotalBuiltArea': ['Built Area'],
    'roadFacing': ['Road Facing'],
    'noOfHouses': ['Houses'],
    'rentalUnits': ['Rentals'],
    'shopsUnits': ['Shops'],
    'parkingUnits': ['Parking'],
    'archLift': ['Lift'],
    'archVastu': ['Vastu'],
    'rooms': ['Rooms'],
    'intFloorSelect': ['Interior Floors'],
    'archSubService': ['Sub Services'],
    'archSubTotal': ['Sub Total'],
    'archGst': ['GST'],
    'archTotal': ['Total'],
    'intTotalBuiltArea': ['Interior Area'],
    'constructionStage': ['Construction Stage'],
    'intSubTotal': ['Interior Subtotal'],
    'intGst': ['Interior GST'],
    'intTotal': ['Interior Total Cost'],
    'floorReq0': ['Ground floor requirement'],
    'floorReq1': ['Floor 1 requirement'],
    'floorReq2': ['Floor 2 requirement'],
    'floorReq3': ['Floor 3 requirement'],
    'floorReq4': ['Floor 4 requirement'],
    'floorReq5': ['Floor 5 requirement']
};

async function getValidAccessToken() {
    let tokens = null;

    // 1. Try to load from tokens.json
    if (fs.existsSync(TOKENS_FILE)) {
        try {
            tokens = JSON.parse(fs.readFileSync(TOKENS_FILE, 'utf8'));
        } catch (err) {
            console.error('Error reading tokens.json:', err.message);
        }
    }

    // 2. If no tokens but we have ZOHO_REFRESH_TOKEN env var, use it
    if (!tokens && ZOHO_REFRESH_TOKEN) {
        console.log('Using ZOHO_REFRESH_TOKEN from environment...');
        tokens = { refresh_token: ZOHO_REFRESH_TOKEN, expires_at: 0 };
    }

    // 3. If still no tokens, try grant code
    if (!tokens) {
        if (!ZOHO_GRANT_CODE) {
            throw new Error('No authentication method available. Set ZOHO_REFRESH_TOKEN or ZOHO_GRANT_CODE.');
        }
        console.log(`Exchanging grant code for tokens...`);
        try {
            const response = await axios.post(`https://accounts.zoho.${ZOHO_REGION}/oauth/v2/token`, null, {
                params: {
                    code: ZOHO_GRANT_CODE,
                    client_id: ZOHO_CLIENT_ID,
                    client_secret: ZOHO_CLIENT_SECRET,
                    grant_type: 'authorization_code'
                }
            });
            tokens = {
                access_token: response.data.access_token,
                refresh_token: response.data.refresh_token,
                expires_at: Date.now() + (response.data.expires_in * 1000)
            };
            try {
                fs.writeFileSync(TOKENS_FILE, JSON.stringify(tokens, null, 2));
            } catch (e) {
                console.warn('Could not write tokens.json (likely read-only cloud environment)');
            }
            return tokens.access_token;
        } catch (err) {
            console.error('Error exchanging grant code:', err.response?.data || err.message);
            throw new Error('Zoho Authentication failed. Please check ZOHO_GRANT_CODE.');
        }
    }

    // 4. Refresh if expired or almost expired
    if (Date.now() >= (tokens.expires_at || 0) - 60000) {
        console.log('Refreshing access token...');
        try {
            const response = await axios.post(`https://accounts.zoho.${ZOHO_REGION}/oauth/v2/token`, null, {
                params: {
                    refresh_token: tokens.refresh_token,
                    client_id: ZOHO_CLIENT_ID,
                    client_secret: ZOHO_CLIENT_SECRET,
                    grant_type: 'refresh_token'
                }
            });
            tokens.access_token = response.data.access_token;
            tokens.expires_at = Date.now() + (response.data.expires_in * 1000);
            try {
                fs.writeFileSync(TOKENS_FILE, JSON.stringify(tokens, null, 2));
            } catch (e) {
                console.warn('Could not update tokens.json (likely read-only cloud environment)');
            }
        } catch (err) {
            console.error('Error refreshing token:', err.response?.data || err.message);
            // If it fails and we have a grant code, maybe it was a one-time thing or the refresh token died
            if (ZOHO_GRANT_CODE) {
                console.log('Refresh failed, but grant code exists. Will try grant code on next run.');
                if (fs.existsSync(TOKENS_FILE)) fs.unlinkSync(TOKENS_FILE);
            }
            throw new Error('Zoho token refresh failed.');
        }
    }
    return tokens.access_token;
}

function getBestMatch(data, header) {
    const hLower = header.toLowerCase().trim().replace(/[^a-z0-9]/g, '');
    if (data[header] !== undefined) return data[header];
    for (let key in data) {
        if (key.toLowerCase().trim().replace(/[^a-z0-9]/g, '') === hLower) return data[key];
    }
    for (let [dataKey, aliases] of Object.entries(FIELD_MAPPINGS)) {
        if (aliases.some(a => a.toLowerCase().replace(/[^a-z0-9]/g, '') === hLower) && data[dataKey] !== undefined) {
            return data[dataKey];
        }
    }
    return undefined;
}

async function appendToZohoSheet(data) {
    try {
        const accessToken = await getValidAccessToken();
        
        // Define all fields in order
        const headers = [
            "Date", "Quotation No", "Full Name", "Contact No", "Mail Id", "Location", "Full Address", 
            "Service", "Online Service", "Budget", "Plot Length", "Plot Width", "Plot Area", 
            "Setback", "Typical Area", "Floors", "Terrace", "Built Area", "Road Facing", 
            "Houses", "Rentals", "Shops", "Parking", "Lift", "Vastu", "Rooms", 
            "Interior Floors", "Sub Services", "Sub Total", "GST", "Total", 
            "Interior Area", "Construction Stage", "Interior Subtotal", "Interior GST", 
            "Interior Total Cost", "Ground floor requirement", "Floor 1 requirement", 
            "Floor 2 requirement", "Floor 3 requirement", "Floor 4 requirement", "Floor 5 requirement"
        ];

        const mappedValues = [
            data.timestamp || new Date().toISOString(),
            data.quotationNo || '',
            data.name || '',
            data.phone || '',
            data.email || '',
            data.city || '',
            data.address || '',
            data.mainService || '',
            data.onlineService || '',
            data.budget || '',
            data.plotLength || '',
            data.plotWidth || '',
            data.archPlotArea || '',
            data.archSetback || '',
            data.typicalAreaAfterSetback || '',
            data.archFloors || '',
            data.archTerrace || '',
            data.archTotalBuiltArea || '',
            data.roadFacing || '',
            data.noOfHouses || '',
            data.rentalUnits || '',
            data.shopsUnits || '',
            data.parkingUnits || '',
            data.archLift || '',
            data.archVastu || '',
            data.rooms || '',
            data.intFloorSelect || '',
            data.archSubService || '',
            data.archSubTotal || '',
            data.archGst || '',
            data.archTotal || '',
            data.intTotalBuiltArea || '',
            data.constructionStage || '',
            data.intSubTotal || '',
            data.intGst || '',
            data.intTotal || '',
            data.floorReq0 || '',
            data.floorReq1 || '',
            data.floorReq2 || '',
            data.floorReq3 || '',
            data.floorReq4 || '',
            data.floorReq5 || ''
        ];
        
        // Check if we need to add headers (search first 10 rows for "Date")
        const checkUrl = `https://sheet.zoho.${ZOHO_REGION}/api/v2/${ZOHO_WORKBOOK_ID}?method=worksheet.content.get&worksheet_name=${ZOHO_SHEET_NAME}&range=A1:A10`;
        const checkRes = await axios.get(checkUrl, {
            headers: { 'Authorization': `Zoho-oauthtoken ${accessToken}` }
        });
        
        let headerFound = false;
        if (checkRes.data.range_details && checkRes.data.range_details.length > 0) {
            // Zoho V2 returns range_details as an array of row objects
            for (const row of checkRes.data.range_details) {
                if (row.row_details && row.row_details.length > 0) {
                    // row_details is an array of cell objects for that row
                    const firstCell = row.row_details[0]?.content;
                    if (firstCell === 'Date') {
                        headerFound = true;
                        break;
                    }
                }
            }
        }
        
        const needsHeaders = !headerFound;
        
        let csvContent = '';
        if (needsHeaders) {
            console.log('Header "Date" not found in first 10 rows, adding headers...');
            csvContent += headers.map(h => `"${String(h).replace(/"/g, '""')}"`).join(',') + '\n';
        }
        csvContent += mappedValues.map(v => `"${String(v).replace(/"/g, '""')}"`).join(',');

        console.log('Appending data via CSV method (works reliably without pre-existing headers)');
        
        const url = `https://sheet.zoho.${ZOHO_REGION}/api/v2/${ZOHO_WORKBOOK_ID}?method=worksheet.csvdata.append`;
        const params = new URLSearchParams();
        params.append('worksheet_name', ZOHO_SHEET_NAME);
        params.append('data', csvContent);

        const response = await axios.post(url, params.toString(), {
            headers: { 
                'Authorization': `Zoho-oauthtoken ${accessToken}`, 
                'Content-Type': 'application/x-www-form-urlencoded' 
            }
        });

        console.log('Zoho response:', response.data.status);
        return response.data;
    } catch (error) {
        console.error('Zoho Error:', error.response?.data || error.message);
        return { status: 'failure', error: error.message };
    }
}

const server = http.createServer(async (req, res) => {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    if (req.method === 'OPTIONS') { res.writeHead(204); res.end(); return; }

    if (req.method === 'POST' && (req.url === '/save' || req.url === '/')) {
        let body = '';
        req.on('data', chunk => body += chunk);
        req.on('end', async () => {
            try {
                if (!body) throw new Error('Empty request body');
                const data = JSON.parse(body);
                console.log('Received data to save:', data.name || 'No name');

                // CSV Save
                const headers = Object.keys(data);
                const fileExists = fs.existsSync(CSV_FILE);
                const csvRow = (fileExists ? '' : headers.join(',') + '\n') + headers.map(h => `"${String(data[h] || '').replace(/"/g, '""')}"`).join(',') + '\n';
                fs.appendFileSync(CSV_FILE, csvRow);

                // Zoho Save
                const zohoRes = await appendToZohoSheet(data);
                
                res.writeHead(200, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ success: true, zoho: zohoRes.status }));
            } catch (e) {
                console.error('Save error:', e.message);
                res.writeHead(400); res.end(JSON.stringify({ success: false, error: e.message }));
            }
        });
    } else { res.writeHead(404); res.end(); }
});

server.listen(PORT, () => console.log(`Backend server active on ${PORT}`));
