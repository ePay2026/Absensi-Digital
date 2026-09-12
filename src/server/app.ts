import express, { Request, Response, NextFunction } from 'express';
import cors from 'cors';
import fs from 'fs';
import path from 'path';
import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';
import { Resend } from 'resend';

// Helper to sanitize and format Google private key for Vercel and container environments
export function formatPrivateKey(rawKey?: string): string {
  if (!rawKey) return '';
  let key = rawKey.trim();

  // If user accidentally pasted the whole service account JSON
  if (key.startsWith('{') && key.endsWith('}')) {
    try {
      const parsed = JSON.parse(key);
      if (parsed.private_key) {
        key = parsed.private_key;
      }
    } catch (e) {}
  }

  // Strip enclosing quotes (single, double, or backticks, even if multiple)
  while (
    (key.startsWith('"') && key.endsWith('"')) ||
    (key.startsWith("'") && key.endsWith("'")) ||
    (key.startsWith('`') && key.endsWith('`'))
  ) {
    key = key.slice(1, -1).trim();
  }

  // Replace literal escaped backslash-n or carriage returns
  key = key.replace(/\\\\r\\\\n/g, '\n').replace(/\\\\n/g, '\n');
  key = key.replace(/\\r\\n/g, '\n').replace(/\\n/g, '\n');
  key = key.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

  // Handle case where BEGIN/END tags are missing newlines
  if (key.includes('-----BEGIN PRIVATE KEY-----')) {
    const beginMarker = '-----BEGIN PRIVATE KEY-----';
    const endMarker = '-----END PRIVATE KEY-----';
    const beginIdx = key.indexOf(beginMarker);
    const endIdx = key.indexOf(endMarker);
    if (beginIdx !== -1 && endIdx !== -1) {
      const header = beginMarker;
      const footer = endMarker;
      let body = key.substring(beginIdx + beginMarker.length, endIdx).trim();
      body = body.replace(/ /g, '\n').replace(/\n+/g, '\n');
      key = header + '\n' + body + '\n' + footer + '\n';
    }
  }

  return key;
}

export function getServiceAccountEmail(): string {
  let email = (process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || '').trim();
  if (email.startsWith('{') && email.endsWith('}')) {
    try {
      const parsed = JSON.parse(email);
      if (parsed.client_email) return parsed.client_email.trim();
    } catch (e) {}
  }
  while ((email.startsWith('"') && email.endsWith('"')) || (email.startsWith("'") && email.endsWith("'"))) {
    email = email.slice(1, -1).trim();
  }
  if (!email && process.env.GOOGLE_PRIVATE_KEY) {
    try {
      const raw = process.env.GOOGLE_PRIVATE_KEY.trim();
      if (raw.startsWith('{') && raw.endsWith('}')) {
        const parsed = JSON.parse(raw);
        if (parsed.client_email) return parsed.client_email.trim();
      }
    } catch (e) {}
  }
  return email;
}

export function getSpreadsheetId(): string {
  let id = (process.env.SPREADSHEET_ID || '').trim();
  while ((id.startsWith('"') && id.endsWith('"')) || (id.startsWith("'") && id.endsWith("'"))) {
    id = id.slice(1, -1).trim();
  }
  const match = id.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match && match[1]) {
    return match[1];
  }
  return id;
}

// Resend Email client getter (lazy evaluation for serverless)
function getResendClient(): Resend | null {
  const apiKey = process.env.RESEND_API_KEY;
  if (!apiKey) return null;
  try {
    return new Resend(apiKey);
  } catch (err) {
    console.error('Failed to initialize Resend client:', err);
    return null;
  }
}

// Fallback In-Memory / File Database
interface MockDB {
  users: any[];
  employees: any[];
  attendance: any[];
  locations: any[];
  settings: any;
  units?: any[];
  admins?: any[];
  shifts?: any[];
  announcements?: any[];
  holidays?: any[];
  deviceBindings?: any[];
}

const defaultDB: MockDB = {
  users: [
    { id: 1, nip: '123456', name: 'Admin User', email: 'admin@puskesmas.com', role: 'admin', password: 'password', office: 'Kantor Induk', group: 'Superadmin' },
    { id: 2, nip: '654321', name: 'Regular User', email: 'user@puskesmas.com', role: 'user', password: 'password', office: 'Pustu A' },
  ],
  employees: [
    { id: '1', name: 'Admin User', nip: '123456', office: 'Kantor Induk', email: 'admin@puskesmas.com', gender: 'Laki-laki', cluster: 'Klaster 1', unit: 'Manajemen' },
    { id: '2', name: 'Regular User', nip: '654321', office: 'Pustu A', email: 'user@puskesmas.com', gender: 'Perempuan', cluster: 'Klaster 2', unit: 'Pustu' }
  ],
  attendance: [],
  locations: [
    { id: 1, name: 'Kantor Induk', desa: 'Kantor Induk', kecamatan: '', kabupaten: '', coordinates: '-7.250445, 112.768845', radius: 300 },
    { id: 2, name: 'Pustu A', desa: 'Pustu A', kecamatan: '', kabupaten: '', coordinates: '-7.1235, 112.1235', radius: 250 },
  ],
  settings: {
    appName: 'Absensi Digital',
    companyName: 'Puskesmas Sehat',
    headName: 'Dr. Budi Santoso',
    address: 'Jl. Kesehatan No. 1, Kota Sehat',
    mainLocation: '-7.250445, 112.768845',
    tolerance: 25,
  },
  units: [
    { id: '1', name: 'Manajemen' },
    { id: '2', name: 'Pustu' },
    { id: '3', name: 'Pelayanan' }
  ],
  shifts: [
    { id: '1', name: 'Pagi', startTime: '08:00', endTime: '16:00', fridayEndTime: '10:50', saturdayEndTime: '12:30', checkInBeforeMinutes: 60, checkInAfterMinutes: 15, checkOutBeforeMinutes: 10, checkOutAfterMinutes: 120, crossesMidnight: false, isActive: true, unit: '' },
    { id: '2', name: 'Malam', startTime: '20:00', endTime: '04:00', fridayEndTime: '', saturdayEndTime: '', checkInBeforeMinutes: 60, checkInAfterMinutes: 15, checkOutBeforeMinutes: 10, checkOutAfterMinutes: 120, crossesMidnight: true, isActive: true, unit: '' }
  ],
  announcements: [],
  holidays: [],
  deviceBindings: []
};

// Determine writable directory for fallback persistence
const isVercel = process.env.VERCEL === '1' || !!process.env.AWS_LAMBDA_FUNCTION_NAME;
const fallbackDbPath = isVercel ? '/tmp/absensi_db.json' : path.join(process.cwd(), 'absensi_db.json');

function loadFallbackDB(): MockDB {
  try {
    if (fs.existsSync(fallbackDbPath)) {
      const data = fs.readFileSync(fallbackDbPath, 'utf-8');
      return { ...defaultDB, ...JSON.parse(data) };
    }
  } catch (e) {
    console.warn('Could not read fallback DB from file:', e);
  }
  return { ...defaultDB };
}

const db: MockDB = loadFallbackDB();

function persistFallbackDB(): void {
  try {
    fs.writeFileSync(fallbackDbPath, JSON.stringify(db, null, 2), 'utf-8');
  } catch (e) {
    // Non-fatal if filesystem is restricted
  }
}

// Google Spreadsheet connection management (Serverless-optimized with singleton memoization)
let docInstance: GoogleSpreadsheet | null = null;
let docInitPromise: Promise<GoogleSpreadsheet | null> | null = null;
let lastSpreadsheetError: string | null = null;
let lastSpreadsheetSuccessTime: number = 0;
const syncedSheets = new Set<string>();

async function seedInitialAdminIfNeeded(doc: GoogleSpreadsheet) {
  try {
    let sheet = doc.sheetsByTitle['Admins'];
    if (!sheet) {
      sheet = await doc.addSheet({
        title: 'Admins',
        headerValues: ['id', 'name', 'nip', 'email', 'phone', 'group', 'isActive', 'access', 'password']
      });
    }
    await sheet.loadHeaderRow();
    const rows = await sheet.getRows();
    if (rows.length === 0) {
      console.log('Seeding initial Superadmin to Google Spreadsheet Admins sheet...');
      await sheet.addRow({
        id: '1',
        name: 'Admin User',
        nip: '123456',
        email: 'admin@puskesmas.com',
        phone: '08123456789',
        group: 'Superadmin',
        isActive: 'true',
        access: JSON.stringify(['Dashboard', 'Absensi', 'Master Data', 'Sistem']),
        password: 'password'
      });
      console.log('Initial Superadmin seeded successfully to Google Spreadsheet');
    }
  } catch (e) {
    console.warn('Auto-seed admin warning:', e);
  }
}

export async function getGoogleDoc(forceFresh = false): Promise<GoogleSpreadsheet | null> {
  if (forceFresh) {
    docInstance = null;
    docInitPromise = null;
    syncedSheets.clear();
  } else {
    if (docInstance) return docInstance;
    if (docInitPromise) return docInitPromise;
  }

  const spreadsheetId = getSpreadsheetId();
  const clientEmail = getServiceAccountEmail();
  const privateKey = formatPrivateKey(process.env.GOOGLE_PRIVATE_KEY);

  if (!spreadsheetId || !clientEmail || !privateKey) {
    const missing: string[] = [];
    if (!spreadsheetId) missing.push('SPREADSHEET_ID');
    if (!clientEmail) missing.push('GOOGLE_SERVICE_ACCOUNT_EMAIL');
    if (!privateKey) missing.push('GOOGLE_PRIVATE_KEY');
    lastSpreadsheetError = `Variabel lingkungan belum lengkap: ${missing.join(', ')}`;
    return null;
  }

  docInitPromise = (async () => {
    try {
      console.log(`Connecting to Google Sheets ID: ${spreadsheetId.substring(0, 8)}... with service email: ${clientEmail}`);
      const serviceAccountAuth = new JWT({
        email: clientEmail,
        key: privateKey,
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
      });
      const spreadsheet = new GoogleSpreadsheet(spreadsheetId, serviceAccountAuth);
      await spreadsheet.loadInfo();
      console.log('Google Spreadsheet connected successfully:', spreadsheet.title);
      docInstance = spreadsheet;
      lastSpreadsheetSuccessTime = Date.now();
      lastSpreadsheetError = null;

      // Seed initial admin if Admins sheet is empty
      seedInitialAdminIfNeeded(spreadsheet).catch(err => {
        console.warn('Initial admin seed task:', err);
      });

      return docInstance;
    } catch (error) {
      const errMsg = error instanceof Error ? error.message : String(error);
      console.error('Failed to connect to Google Spreadsheet:', errMsg);
      lastSpreadsheetError = errMsg;
      docInstance = null;
      return null;
    } finally {
      docInitPromise = null;
    }
  })();

  return docInitPromise;
}

// In-Memory cache for read performance (15-second TTL for fast updates)
const cache: { [key: string]: { data: any; timestamp: number } } = {};
const inFlightRequests: { [key: string]: Promise<any> } = {};
const CACHE_DURATION = 15 * 1000; // 15 seconds cache

async function getCachedData(key: string, fetchFn: () => Promise<any>, bypassCache = false) {
  if (!bypassCache && cache[key] && Date.now() - cache[key].timestamp < CACHE_DURATION) {
    return cache[key].data;
  }
  if (!bypassCache && inFlightRequests[key]) {
    return inFlightRequests[key];
  }
  const promise = fetchFn()
    .then(data => {
      cache[key] = { data, timestamp: Date.now() };
      return data;
    })
    .finally(() => {
      delete inFlightRequests[key];
    });
  inFlightRequests[key] = promise;
  return promise;
}

// Sheet retrieval helper
async function getSheet(title: string) {
  const doc = await getGoogleDoc();
  if (!doc) return null;
  try {
    const sheet = doc.sheetsByTitle[title];
    if (!sheet) {
      console.warn(`Sheet '${title}' not found in spreadsheet.`);
      return null;
    }
    return sheet;
  } catch (error) {
    console.error(`Error getting sheet ${title}:`, error);
    return null;
  }
}

// Sheet creation or header synchronization helper
async function getOrCreateSheet(title: string, headerValues: string[]) {
  const doc = await getGoogleDoc();
  if (!doc) return null;
  let sheet = doc.sheetsByTitle[title];
  if (!sheet) {
    try {
      sheet = await doc.addSheet({ title, headerValues });
      syncedSheets.add(title);
    } catch (e) {
      console.error(`Error creating sheet '${title}':`, e);
      return null;
    }
  } else if (!syncedSheets.has(title)) {
    try {
      await sheet.loadHeaderRow();
      const currentHeaders = sheet.headerValues || [];
      let headersChanged = false;
      const newHeaders = [...currentHeaders];
      for (const header of headerValues) {
        if (!newHeaders.includes(header)) {
          newHeaders.push(header);
          headersChanged = true;
        }
      }
      if (headersChanged) {
        await sheet.setHeaderRow(newHeaders);
      }
      syncedSheets.add(title);
    } catch (e) {
      try {
        await sheet.setHeaderRow(headerValues);
        syncedSheets.add(title);
      } catch (err) {
        console.error(`Failed setting headers for sheet '${title}':`, err);
      }
    }
  }
  return sheet;
}

// Create Express Application and Router
const app = express();
const apiRouter = express.Router();

// Middlewares
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));

// ----------------------------------------------------
// API ROUTES IMPLEMENTATION
// ----------------------------------------------------

// --- Health Check ---
apiRouter.get('/healthz', async (_req: Request, res: Response) => {
  const doc = await getGoogleDoc();
  res.status(200).json({
    status: 'ok',
    timestamp: Date.now(),
    vercel: isVercel,
    spreadsheetConnected: !!doc,
    spreadsheetTitle: doc?.title || null
  });
});

// --- Google Spreadsheet Diagnostic & Connection Test ---
apiRouter.get('/spreadsheet-status', async (req: Request, res: Response) => {
  const forceFresh = req.query.test === 'true';
  const doc = await getGoogleDoc(forceFresh);
  const spreadsheetId = getSpreadsheetId();
  const clientEmail = getServiceAccountEmail();
  const rawKey = process.env.GOOGLE_PRIVATE_KEY || '';
  const formattedKey = formatPrivateKey(rawKey);

  const envCheck = {
    hasSpreadsheetId: !!spreadsheetId,
    spreadsheetIdPreview: spreadsheetId ? `${spreadsheetId.substring(0, 8)}...${spreadsheetId.substring(Math.max(0, spreadsheetId.length - 6))}` : null,
    hasClientEmail: !!clientEmail,
    clientEmailPreview: clientEmail || null,
    hasPrivateKey: !!rawKey,
    privateKeyLength: rawKey.length,
    privateKeyValidPEM: formattedKey.includes('-----BEGIN PRIVATE KEY-----') && formattedKey.includes('-----END PRIVATE KEY-----'),
    isVercel
  };

  if (!doc) {
    return res.json({
      connected: false,
      message: 'Belum dapat terhubung ke Google Spreadsheet',
      error: lastSpreadsheetError || 'Koneksi gagal. Pastikan email service account sudah diberi akses Editor di Google Sheets dan private key valid.',
      envCheck
    });
  }

  const sheetStats: Record<string, { title: string; rowCount: number }> = {};
  for (const [title, sheet] of Object.entries(doc.sheetsByTitle)) {
    sheetStats[title] = {
      title,
      rowCount: sheet.rowCount
    };
  }

  res.json({
    connected: true,
    message: 'Berhasil terhubung ke Google Spreadsheet',
    spreadsheetTitle: doc.title,
    spreadsheetId,
    clientEmail,
    lastConnected: lastSpreadsheetSuccessTime,
    sheetCount: Object.keys(doc.sheetsByTitle).length,
    sheets: Object.keys(doc.sheetsByTitle),
    sheetStats,
    envCheck
  });
});

apiRouter.post('/spreadsheet-status', async (_req: Request, res: Response) => {
  // Clear all in-memory caches
  Object.keys(cache).forEach(k => delete cache[k]);
  const doc = await getGoogleDoc(true);
  if (!doc) {
    return res.status(500).json({
      connected: false,
      message: 'Gagal menghubungkan ke Google Spreadsheet',
      error: lastSpreadsheetError
    });
  }
  return res.json({
    connected: true,
    message: `Koneksi berhasil diverifikasi! Tersambung ke Google Spreadsheet "${doc.title}"`,
    spreadsheetTitle: doc.title,
    sheetCount: Object.keys(doc.sheetsByTitle).length,
    sheets: Object.keys(doc.sheetsByTitle)
  });
});

// --- Server Time Sync ---
apiRouter.get('/time', (_req: Request, res: Response) => {
  res.json({ timestamp: Date.now() });
});

// --- Employees API ---
apiRouter.get('/employees', async (req: Request, res: Response) => {
  const forceFresh = req.query.fresh === 'true';
  const doc = await getGoogleDoc(forceFresh);
  if (doc) {
    try {
      const employees = await getCachedData('employees', async () => {
        const sheet = await getOrCreateSheet('Employees', ['id', 'name', 'nip', 'office', 'office2', 'email', 'gender', 'cluster', 'unit', 'password', 'photoUrl', 'photoUploadCount']);
        if (sheet) {
          const rows = await sheet.getRows();
          return rows.map(row => ({
            id: row.get('id'),
            name: row.get('name'),
            nip: row.get('nip'),
            office: row.get('office'),
            office2: row.get('office2'),
            email: row.get('email'),
            gender: row.get('gender'),
            cluster: row.get('cluster'),
            unit: row.get('unit'),
            password: row.get('password'),
            photoUrl: row.get('photoUrl'),
            photoUploadCount: row.get('photoUploadCount') ? parseInt(row.get('photoUploadCount'), 10) : 0
          }));
        }
        return [];
      }, forceFresh);
      return res.json(employees);
    } catch (error) {
      console.error('Error fetching employees from spreadsheet:', error);
    }
  }
  res.json(db.employees);
});

apiRouter.post('/employees', async (req: Request, res: Response) => {
  const employee = req.body;
  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const sheet = await getOrCreateSheet('Employees', ['id', 'name', 'nip', 'office', 'office2', 'email', 'gender', 'cluster', 'unit', 'password', 'photoUrl', 'photoUploadCount']);
      if (sheet) {
        await sheet.addRow({
          ...employee,
          office2: employee.office2 || ''
        });
        delete cache['employees'];
        return res.json({ success: true, message: 'Karyawan berhasil ditambahkan' });
      }
    } catch (error) {
      console.error('Error saving employee to spreadsheet:', error);
      return res.status(500).json({ success: false, message: 'Gagal menyimpan karyawan ke spreadsheet' });
    }
  }

  db.employees.push(employee);
  persistFallbackDB();
  res.json({ success: true, message: 'Karyawan berhasil ditambahkan' });
});

apiRouter.post('/employees/bulk', async (req: Request, res: Response) => {
  const employeesData = req.body;
  if (!Array.isArray(employeesData)) {
    return res.status(400).json({ success: false, message: 'Data harus berupa array' });
  }

  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const sheet = await getOrCreateSheet('Employees', ['id', 'name', 'nip', 'office', 'office2', 'email', 'gender', 'cluster', 'unit', 'password', 'photoUrl', 'photoUploadCount']);
      if (sheet) {
        const rows = employeesData.map(emp => ({
          ...emp,
          office2: emp.office2 || ''
        }));
        await sheet.addRows(rows);
        delete cache['employees'];
        return res.json({ success: true, message: `${employeesData.length} karyawan berhasil ditambahkan` });
      }
    } catch (error) {
      console.error('Error saving bulk employees to spreadsheet:', error);
      return res.status(500).json({ success: false, message: `Gagal menyimpan data ke spreadsheet: ${error instanceof Error ? error.message : String(error)}` });
    }
  }

  db.employees.push(...employeesData);
  persistFallbackDB();
  res.json({ success: true, message: `${employeesData.length} karyawan berhasil ditambahkan` });
});

apiRouter.delete('/employees/:id', async (req: Request, res: Response) => {
  const { id } = req.params;
  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const sheet = await getSheet('Employees');
      if (sheet) {
        const rows = await sheet.getRows();
        const rowToDelete = rows.find(r => r.get('id') === id);
        if (rowToDelete) {
          await rowToDelete.delete();
          delete cache['employees'];
          return res.json({ success: true, message: 'Karyawan berhasil dihapus' });
        }
      }
    } catch (error) {
      console.error('Error deleting employee from spreadsheet:', error);
      return res.status(500).json({ success: false, message: 'Gagal menghapus karyawan dari spreadsheet' });
    }
  }

  db.employees = db.employees.filter(e => e.id !== id);
  persistFallbackDB();
  res.json({ success: true, message: 'Karyawan berhasil dihapus' });
});

apiRouter.post('/employees/photo', async (req: Request, res: Response) => {
  const { nip, photoUrl } = req.body;
  if (!nip || !photoUrl) {
    return res.status(400).json({ success: false, message: 'Data tidak lengkap' });
  }

  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const sheet = await getSheet('Employees');
      if (sheet) {
        const rows = await sheet.getRows();
        const empRow = rows.find(r => String(r.get('nip')) === String(nip));
        if (empRow) {
          const currentCount = empRow.get('photoUploadCount') ? parseInt(empRow.get('photoUploadCount'), 10) : 0;
          if (currentCount >= 5) {
            return res.status(400).json({ success: false, message: 'Batas unggah foto telah mencapai maksimal (5 kali).' });
          }
          empRow.set('photoUrl', photoUrl);
          empRow.set('photoUploadCount', String(currentCount + 1));
          await empRow.save();
          delete cache['employees'];
          return res.json({
            success: true,
            message: `Foto berhasil disimpan. Sisa kesempatan: ${4 - currentCount} kali.`,
            photoUrl,
            photoUploadCount: currentCount + 1
          });
        } else {
          return res.status(404).json({ success: false, message: 'Karyawan tidak ditemukan' });
        }
      }
    } catch (error) {
      console.error('Error updating profile photo:', error);
      return res.status(500).json({ success: false, message: 'Terjadi kesalahan sistem saat menyimpan foto.' });
    }
  }

  const emp = db.employees.find((e: any) => e.nip === nip) as any;
  if (emp) {
    const currentCount = emp.photoUploadCount || 0;
    if (currentCount >= 5) {
      return res.status(400).json({ success: false, message: 'Batas unggah foto telah mencapai maksimal (5 kali).' });
    }
    emp.photoUrl = photoUrl;
    emp.photoUploadCount = currentCount + 1;
    persistFallbackDB();
    return res.json({
      success: true,
      message: `Foto berhasil disimpan. Sisa kesempatan: ${4 - currentCount} kali.`,
      photoUrl,
      photoUploadCount: currentCount + 1
    });
  }

  res.status(404).json({ success: false, message: 'Karyawan tidak ditemukan' });
});

// --- Admins API ---
apiRouter.get('/admins', async (req: Request, res: Response) => {
  const forceFresh = req.query.fresh === 'true';
  const doc = await getGoogleDoc(forceFresh);
  if (doc) {
    try {
      let admins = await getCachedData('admins', async () => {
        const sheet = await getOrCreateSheet('Admins', ['id', 'name', 'nip', 'email', 'phone', 'group', 'isActive', 'access', 'password']);
        if (sheet) {
          const rows = await sheet.getRows();
          return rows.map(row => ({
            id: row.get('id'),
            name: row.get('name'),
            nip: row.get('nip'),
            email: row.get('email'),
            phone: row.get('phone'),
            group: row.get('group'),
            isActive: String(row.get('isActive')).toLowerCase() === 'true',
            access: row.get('access') ? JSON.parse(row.get('access')) : [],
            password: row.get('password')
          }));
        }
        return [];
      }, forceFresh);

      if (admins.length === 0) {
        const sheet = await getSheet('Admins');
        if (sheet) {
          const defaultAdmin = {
            id: '1',
            name: 'Admin User',
            nip: '123456',
            email: 'admin@puskesmas.com',
            phone: '08123456789',
            group: 'Superadmin',
            isActive: 'true',
            access: JSON.stringify(['Dashboard', 'Absensi', 'Master Data', 'Sistem']),
            password: 'password'
          };
          await sheet.addRow(defaultAdmin);
          delete cache['admins'];
          admins = [{
            ...defaultAdmin,
            isActive: true,
            access: ['Dashboard', 'Absensi', 'Master Data', 'Sistem']
          }];
        }
      }

      return res.json(admins);
    } catch (error) {
      console.error('Error fetching admins from spreadsheet:', error);
    }
  }
  res.json(db.admins || []);
});

apiRouter.post('/admins', async (req: Request, res: Response) => {
  const admin = req.body;
  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const sheet = await getOrCreateSheet('Admins', ['id', 'name', 'nip', 'email', 'phone', 'group', 'isActive', 'access', 'password']);
      if (sheet) {
        await sheet.addRow({
          ...admin,
          isActive: (admin.isActive ?? true).toString(),
          access: JSON.stringify(admin.access || [])
        });
        delete cache['admins'];
        return res.json({ success: true, message: 'Admin berhasil ditambahkan' });
      }
    } catch (error) {
      console.error('Error saving admin to spreadsheet:', error);
      return res.status(500).json({ success: false, message: 'Gagal menyimpan admin ke spreadsheet' });
    }
  }

  if (!db.admins) db.admins = [];
  db.admins.push(admin);
  persistFallbackDB();
  res.json({ success: true, message: 'Admin berhasil ditambahkan' });
});

apiRouter.delete('/admins/:id', async (req: Request, res: Response) => {
  const { id } = req.params;
  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const sheet = await getSheet('Admins');
      if (sheet) {
        const rows = await sheet.getRows();
        const rowToDelete = rows.find(r => String(r.get('id')) === String(id));
        if (rowToDelete) {
          await rowToDelete.delete();
          delete cache['admins'];
          return res.json({ success: true, message: 'Admin berhasil dihapus' });
        }
      }
    } catch (error) {
      console.error('Error deleting admin from spreadsheet:', error);
      return res.status(500).json({ success: false, message: 'Gagal menghapus admin dari spreadsheet' });
    }
  }

  if (db.admins) {
    db.admins = db.admins.filter(a => String(a.id) !== String(id));
    persistFallbackDB();
  }
  res.json({ success: true, message: 'Admin berhasil dihapus' });
});

apiRouter.put('/admins/:id', async (req: Request, res: Response) => {
  const { id } = req.params;
  const admin = req.body;
  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const sheet = await getSheet('Admins');
      if (sheet) {
        const rows = await sheet.getRows();
        const rowToUpdate = rows.find(r => String(r.get('id')) === String(id));
        if (rowToUpdate) {
          rowToUpdate.set('name', admin.name);
          rowToUpdate.set('nip', admin.nip);
          rowToUpdate.set('email', admin.email);
          rowToUpdate.set('phone', admin.phone);
          rowToUpdate.set('group', admin.group);
          rowToUpdate.set('isActive', (admin.isActive ?? true).toString());
          rowToUpdate.set('access', JSON.stringify(admin.access || []));
          if (admin.password) rowToUpdate.set('password', admin.password);
          await rowToUpdate.save();
          delete cache['admins'];
          return res.json({ success: true, message: 'Admin berhasil diperbarui' });
        }
      }
    } catch (error) {
      console.error('Error updating admin in spreadsheet:', error);
      return res.status(500).json({ success: false, message: 'Gagal memperbarui admin di spreadsheet' });
    }
  }

  if (db.admins) {
    const idx = db.admins.findIndex(a => String(a.id) === String(id));
    if (idx >= 0) {
      db.admins[idx] = { ...db.admins[idx], ...admin };
      persistFallbackDB();
    }
  }
  res.json({ success: true, message: 'Admin berhasil diperbarui' });
});

// --- Auth API ---
apiRouter.post('/login', async (req: Request, res: Response) => {
  const nip = (req.body.nip || '').trim();
  const password = (req.body.password || '').trim();
  console.log(`Login attempt for NIP: ${nip}`);

  let user: any = null;
  const doc = await getGoogleDoc();
  if (doc) {
    try {
      // 1. Check Admins
      const adminSheet = await getSheet('Admins');
      if (adminSheet) {
        const rows = await adminSheet.getRows();
        const row = rows.find(r =>
          String(r.get('nip') || '').trim() === nip &&
          String(r.get('password') || '').trim() === password &&
          String(r.get('isActive')).trim().toLowerCase() === 'true'
        );
        if (row) {
          let access = [];
          try {
            access = JSON.parse(row.get('access'));
          } catch (e) {}
          user = {
            id: row.get('id'),
            nip: String(row.get('nip') || '').trim(),
            name: row.get('name'),
            role: 'admin',
            group: row.get('group'),
            access
          };
          console.log('Admin authenticated from Admins sheet:', user.name);
        }
      }

      // 2. Check Employees (Official Employees directory)
      if (!user) {
        const empSheet = await getSheet('Employees');
        if (empSheet) {
          const rows = await empSheet.getRows();
          const row = rows.find(r =>
            String(r.get('nip') || '').trim() === nip &&
            String(r.get('password') || '').trim() === password
          );
          if (row) {
            user = {
              id: row.get('id'),
              nip: String(row.get('nip') || '').trim(),
              name: row.get('name'),
              role: 'user',
              office: row.get('office'),
              office2: row.get('office2') || '',
              unit: row.get('unit') || '',
              gender: row.get('gender') || '',
              cluster: row.get('cluster') || ''
            };
            console.log('Employee authenticated from Employees sheet:', user.name);
          }
        }
      }

      // 3. Check Users
      if (!user) {
        const userSheet = await getSheet('Users');
        if (userSheet) {
          const rows = await userSheet.getRows();
          const row = rows.find(r =>
            String(r.get('nip') || '').trim() === nip &&
            String(r.get('password') || '').trim() === password
          );
          if (row) {
            user = {
              id: row.get('id'),
              nip: String(row.get('nip') || '').trim(),
              name: row.get('name'),
              role: row.get('role') || 'user',
              office: row.get('office'),
              office2: row.get('office2') || '',
              unit: row.get('unit') || ''
            };
            console.log('User authenticated from Users sheet:', user.name);
          }
        }
      }
    } catch (error) {
      console.error('Error during spreadsheet login:', error);
    }
  }

  // Fallback Mock DB
  if (!user) {
    user = db.users.find(u => u.nip === nip && u.password === password);
    if (user) console.log('User authenticated via mock DB:', user.name);
  }

  if (user) {
    // Device binding verification for non-admin employees
    if (user.role !== 'admin') {
      const deviceId = req.body.deviceId;
      if (deviceId && doc) {
        try {
          const deviceSheet = await getOrCreateSheet('DeviceBindings', ['nip', 'deviceId']);
          if (deviceSheet) {
            const rows = await deviceSheet.getRows();
            const existingDeviceRow = rows.find(r => r.get('deviceId') === deviceId);
            if (existingDeviceRow && existingDeviceRow.get('nip') !== nip) {
              return res.status(403).json({
                success: false,
                message: 'Perangkat ini sudah digunakan oleh akun lain. Silahkan hubungi Admin untuk mereset perangkat jika fitur ini bermasalah.'
              });
            }

            const existingNipRow = rows.find(r => r.get('nip') === nip);
            if (existingNipRow && existingNipRow.get('deviceId') !== deviceId) {
              return res.status(403).json({
                success: false,
                message: 'Akun Anda terdaftar di perangkat lain. Untuk pengguna iOS/iPhone yang baru menginstall ke Layar Utama, layar utama dianggap sebagai perangkat baru. Silahkan minta Admin untuk mereset perangkat Anda di menu Karyawan.'
              });
            }

            if (!existingDeviceRow && !existingNipRow) {
              await deviceSheet.addRow({ nip, deviceId });
            }
          }
        } catch (error) {
          console.error('Error verifying device binding:', error);
        }
      }
    }
    return res.json({ success: true, user });
  }

  return res.status(401).json({ success: false, message: 'NIP atau Password salah' });
});

apiRouter.post('/change-password', async (req: Request, res: Response) => {
  const { id, role, oldPassword, newPassword } = req.body;
  if (!id || !role || !oldPassword || !newPassword) {
    return res.status(400).json({ success: false, message: 'Data tidak lengkap' });
  }

  let passwordUpdated = false;
  const doc = await getGoogleDoc();
  if (doc) {
    try {
      if (role === 'admin') {
        const sheet = await getSheet('Admins');
        if (sheet) {
          const rows = await sheet.getRows();
          const userRow = rows.find(r => String(r.get('id')) === String(id) && String(r.get('password')) === String(oldPassword));
          if (userRow) {
            userRow.set('password', newPassword);
            await userRow.save();
            delete cache['admins'];
            return res.json({ success: true, message: 'Password admin berhasil diubah' });
          } else {
            return res.status(400).json({ success: false, message: 'Password lama salah' });
          }
        }
      } else {
        const empSheet = await getSheet('Employees');
        let empUpdated = false;
        if (empSheet) {
          const empRows = await empSheet.getRows();
          const empRow = empRows.find(r => (String(r.get('id')) === String(id) || String(r.get('nip')) === String(id)) && String(r.get('password')) === String(oldPassword));
          if (empRow) {
            empRow.set('password', newPassword);
            await empRow.save();
            empUpdated = true;
            delete cache['employees'];
          }
        }

        const userSheet = await getSheet('Users');
        if (userSheet) {
          const userRows = await userSheet.getRows();
          const userRow = userRows.find(r => (String(r.get('id')) === String(id) || String(r.get('nip')) === String(id)) && String(r.get('password')) === String(oldPassword));
          if (userRow) {
            userRow.set('password', newPassword);
            await userRow.save();
            empUpdated = true;
          }
        }

        if (empUpdated) {
          return res.json({ success: true, message: 'Password berhasil diubah' });
        }
      }
    } catch (error) {
      console.error('Error changing password in spreadsheet:', error);
      return res.status(500).json({ success: false, message: 'Terjadi kesalahan sistem' });
    }
  }

  if (!passwordUpdated) {
    const user = db.users.find(u => String(u.id) === String(id));
    if (user) {
      if (user.password === oldPassword) {
        user.password = newPassword;
        if (role !== 'admin' && user.nip) {
          const emp = db.employees.find(e => e.nip === user.nip);
          if (emp) (emp as any).password = newPassword;
        }
        persistFallbackDB();
        return res.json({ success: true, message: 'Password berhasil diubah' });
      } else {
        return res.status(400).json({ success: false, message: 'Password lama salah' });
      }
    } else {
      return res.status(404).json({ success: false, message: 'Pengguna tidak ditemukan' });
    }
  }

  res.json({ success: true, message: 'Password berhasil diubah' });
});

apiRouter.delete('/device-bindings/:nip', async (req: Request, res: Response) => {
  const { nip } = req.params;
  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const deviceSheet = await getSheet('DeviceBindings');
      if (deviceSheet) {
        const rows = await deviceSheet.getRows();
        const existingNipRow = rows.find(r => r.get('nip') === nip);
        if (existingNipRow) {
          await existingNipRow.delete();
          return res.json({ success: true, message: 'Binding perangkat berhasil dihapus.' });
        } else {
          return res.status(404).json({ success: false, message: 'Binding perangkat tidak ditemukan.' });
        }
      }
    } catch (error) {
      console.error('Error resetting device binding:', error);
      return res.status(500).json({ success: false, message: 'Terjadi kesalahan pada server.' });
    }
  }

  if (db.deviceBindings) {
    db.deviceBindings = db.deviceBindings.filter(d => d.nip !== nip);
    persistFallbackDB();
  }
  return res.json({ success: true, message: 'Binding perangkat berhasil dihapus.' });
});

apiRouter.post('/register', async (req: Request, res: Response) => {
  const { nip, name, email, password, gender, cluster, unit, desa, office2 } = req.body;

  let isValidEmployee = false;
  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const empSheet = await getSheet('Employees');
      if (empSheet) {
        const rows = await empSheet.getRows();
        isValidEmployee = rows.some(r => String(r.get('nip')) === String(nip));
      }
    } catch (error) {
      console.error('Error validating employee NIP:', error);
    }
  }

  if (!isValidEmployee) {
    isValidEmployee = db.employees.some(e => e.nip === nip);
  }

  if (!isValidEmployee) {
    return res.status(400).json({ success: false, message: 'NIP tidak terdaftar sebagai karyawan. Hubungi Admin.' });
  }

  let userExists = false;
  if (doc) {
    try {
      const userSheet = await getOrCreateSheet('Users', ['id', 'nip', 'name', 'email', 'role', 'password', 'gender', 'cluster', 'unit', 'office', 'office2']);
      if (userSheet) {
        const rows = await userSheet.getRows();
        userExists = rows.some(r => String(r.get('nip')) === String(nip));
      }
    } catch (error) {
      console.error('Error checking existing user:', error);
    }
  } else {
    userExists = db.users.some(u => u.nip === nip);
  }

  if (userExists) {
    return res.status(400).json({ success: false, message: 'NIP sudah terdaftar sebagai user' });
  }

  const newUser = {
    id: Date.now().toString(),
    nip,
    name,
    email,
    role: 'user',
    password,
    gender,
    cluster,
    unit,
    office: desa,
    office2: office2 || ''
  };

  if (doc) {
    try {
      const sheet = await getOrCreateSheet('Users', ['id', 'nip', 'name', 'email', 'role', 'password', 'gender', 'cluster', 'unit', 'office', 'office2']);
      if (sheet) {
        await sheet.addRow(newUser);
      }
    } catch (error) {
      console.error('Error saving user to spreadsheet:', error);
    }
  } else {
    db.users.push(newUser as any);
    persistFallbackDB();
  }

  res.json({ success: true, message: 'Pendaftaran berhasil' });
});

// --- Attendance API ---
apiRouter.get('/attendance', async (req: Request, res: Response) => {
  const { startDate, endDate } = req.query;
  let allAttendance: any[] = [];
  const doc = await getGoogleDoc();

  if (doc) {
    try {
      allAttendance = await getCachedData('attendance', async () => {
        const sheet = await getOrCreateSheet('Attendance', ['id', 'nip', 'name', 'date', 'time', 'type', 'location', 'status', 'photoUrl', 'shift']);
        if (sheet) {
          const rows = await sheet.getRows();
          return rows.map(row => ({
            id: row.get('id'),
            nip: row.get('nip'),
            name: row.get('name'),
            date: row.get('date'),
            time: row.get('time'),
            type: row.get('type'),
            location: (() => {
              try {
                return JSON.parse(row.get('location'));
              } catch (e) {
                return row.get('location');
              }
            })(),
            status: row.get('status'),
            photoUrl: row.get('photoUrl') && row.get('photoUrl').startsWith('data:image')
              ? `/api/attendance/${row.get('id')}/photo`
              : (row.get('photoUrl') || ''),
            shift: row.get('shift')
          }));
        }
        return [];
      });
    } catch (error) {
      console.error('Error fetching attendance from spreadsheet:', error);
    }
  } else {
    allAttendance = db.attendance;
  }

  const filterByDateRange = (start: string, end: string) => {
    return allAttendance.filter(a => {
      const aStart = a.date;
      const aEnd = (a.location && typeof a.location === 'object' && a.location.endDate) ? a.location.endDate : a.date;
      return aStart <= end && aEnd >= start;
    });
  };

  let filteredAttendance = allAttendance;
  if (startDate && endDate) {
    filteredAttendance = filterByDateRange(startDate as string, endDate as string);
  } else {
    const today = new Date();
    const prevMonthObj = new Date(today.getFullYear(), today.getMonth() - 1, 1);
    const firstDay = `${prevMonthObj.getFullYear()}-${String(prevMonthObj.getMonth() + 1).padStart(2, '0')}-01`;
    const nextMonthObj = new Date(today.getFullYear(), today.getMonth() + 1, 0);
    const nextMonthEnd = `${nextMonthObj.getFullYear()}-${String(nextMonthObj.getMonth() + 1).padStart(2, '0')}-${String(nextMonthObj.getDate()).padStart(2, '0')}`;
    filteredAttendance = filterByDateRange(firstDay, nextMonthEnd);
  }

  res.json(filteredAttendance);
});

apiRouter.get('/attendance/:id/photo', async (req: Request, res: Response) => {
  try {
    const doc = await getGoogleDoc();
    let photoUrl = '';

    if (doc) {
      const sheet = await getOrCreateSheet('Attendance', ['id', 'nip', 'name', 'date', 'time', 'type', 'location', 'status', 'photoUrl', 'shift']);
      if (sheet) {
        const rows = await sheet.getRows();
        const targetRow = rows.find(row => row.get('id') === req.params.id);
        if (targetRow) photoUrl = targetRow.get('photoUrl');
      }
    } else {
      const target = db.attendance.find((a: any) => a.id === req.params.id);
      if (target) photoUrl = target.photoUrl;
    }

    if (!photoUrl || !photoUrl.startsWith('data:image')) {
      return res.status(404).send('No image for this record');
    }

    const matches = photoUrl.match(/^data:([A-Za-z-+\/]+);base64,(.+)$/);
    if (!matches || matches.length !== 3) {
      return res.status(400).send('Invalid image data');
    }

    const mimeType = matches[1];
    const buffer = Buffer.from(matches[2], 'base64');
    res.set('Content-Type', mimeType);
    res.send(buffer);
  } catch (e) {
    console.error('Error sending attendance photo:', e);
    res.status(500).send('Internal error');
  }
});

apiRouter.post('/attendance', async (req: Request, res: Response) => {
  const attendanceData = req.body;

  // Enforce Asia/Jakarta Server Time
  if (attendanceData.type === 'in' || attendanceData.type === 'out') {
    const now = new Date();
    const timeFormatter = new Intl.DateTimeFormat('id-ID', {
      timeZone: 'Asia/Jakarta',
      hour: '2-digit',
      minute: '2-digit',
      hour12: false
    });
    attendanceData.time = timeFormatter.format(now).replace('.', ':');

    if (attendanceData.type === 'in') {
      const dateFormatter = new Intl.DateTimeFormat('en-US', {
        timeZone: 'Asia/Jakarta',
        year: 'numeric',
        month: '2-digit',
        day: '2-digit'
      });
      const parts = dateFormatter.formatToParts(now);
      const year = parts.find(p => p.type === 'year')?.value;
      const month = parts.find(p => p.type === 'month')?.value;
      const day = parts.find(p => p.type === 'day')?.value;
      attendanceData.date = `${year}-${month}-${day}`;
    }
  }

  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const sheet = await getOrCreateSheet('Attendance', ['id', 'nip', 'name', 'date', 'time', 'type', 'location', 'status', 'photoUrl', 'shift']);
      if (sheet) {
        let photoUrlToSave = attendanceData.photoUrl || '';
        if (photoUrlToSave.length > 49000) {
          photoUrlToSave = 'Image too large to save in spreadsheet';
          console.warn('Attendance photoUrl exceeded 50,000 characters, replacing with placeholder.');
        }

        await sheet.addRow({
          id: Date.now().toString(),
          ...attendanceData,
          photoUrl: photoUrlToSave,
          location: typeof attendanceData.location === 'object' ? JSON.stringify(attendanceData.location) : attendanceData.location
        });
        delete cache['attendance'];
        return res.json({ success: true, message: 'Absensi berhasil dicatat' });
      }
    } catch (error) {
      console.error('Error saving attendance to spreadsheet:', error);
      return res.status(500).json({ success: false, message: 'Gagal menyimpan absensi ke spreadsheet. Mungkin ukuran foto terlalu besar.' });
    }
  }

  db.attendance.push({ id: Date.now().toString(), ...attendanceData } as any);
  persistFallbackDB();
  res.json({ success: true, message: 'Absensi berhasil dicatat' });
});

apiRouter.post('/attendance/bulk', async (req: Request, res: Response) => {
  const { records } = req.body;
  if (!Array.isArray(records) || records.length === 0) {
    return res.json({ success: true, message: 'No records to add' });
  }

  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const sheet = await getOrCreateSheet('Attendance', ['id', 'nip', 'name', 'date', 'time', 'type', 'location', 'status', 'photoUrl', 'shift']);
      if (sheet) {
        const rowsToAdd = records.map((attendanceData: any, index: number) => ({
          id: (Date.now() + index).toString(),
          ...attendanceData,
          photoUrl: attendanceData.photoUrl || '',
          location: typeof attendanceData.location === 'object' ? JSON.stringify(attendanceData.location) : attendanceData.location
        }));
        await sheet.addRows(rowsToAdd);
        delete cache['attendance'];
        return res.json({ success: true, message: 'Bulk absensi berhasil dicatat' });
      }
    } catch (error) {
      console.error('Error saving bulk attendance:', error);
      return res.status(500).json({ success: false, message: 'Gagal menyimpan bulk absensi.' });
    }
  }

  records.forEach((rec: any, i: number) => {
    db.attendance.push({ id: (Date.now() + i).toString(), ...rec } as any);
  });
  persistFallbackDB();
  res.json({ success: true, message: 'Bulk absensi berhasil dicatat' });
});

apiRouter.post('/attendance/auto-checkout-check', async (_req: Request, res: Response) => {
  const doc = await getGoogleDoc();
  if (!doc) return res.json({ success: true, fixedCount: 0 });

  try {
    const [attSheet, empSheet, shiftSheet] = await Promise.all([
      getOrCreateSheet('Attendance', ['id', 'nip', 'name', 'date', 'time', 'type', 'location', 'status', 'photoUrl', 'shift']),
      getOrCreateSheet('Employees', ['id', 'name', 'nip', 'office', 'office2', 'email', 'gender', 'cluster', 'unit', 'password', 'photoUrl', 'photoUploadCount']),
      getOrCreateSheet('Shifts', ['id', 'name', 'startTime', 'endTime', 'fridayEndTime', 'saturdayEndTime', 'checkInBeforeMinutes', 'checkInAfterMinutes', 'checkOutBeforeMinutes', 'checkOutAfterMinutes', 'crossesMidnight', 'isActive', 'unit', 'isOffSunday', 'isOffHoliday'])
    ]);

    if (!attSheet || !empSheet || !shiftSheet) return res.json({ success: false });

    const [attRows, empRows, shiftRows] = await Promise.all([
      attSheet.getRows(),
      empSheet.getRows(),
      shiftSheet.getRows()
    ]);

    const attendance = attRows.map(r => ({
      nip: r.get('nip'),
      name: r.get('name'),
      date: r.get('date'),
      time: r.get('time'),
      type: r.get('type'),
      location: r.get('location'),
      status: r.get('status'),
      shift: r.get('shift') || ''
    }));

    const employees = empRows.map(r => ({ nip: r.get('nip'), unit: r.get('unit') }));
    const shifts = shiftRows.map(r => ({
      name: r.get('name'),
      startTime: r.get('startTime'),
      endTime: r.get('endTime'),
      fridayEndTime: r.get('fridayEndTime') || '',
      saturdayEndTime: r.get('saturdayEndTime') || '',
      checkOutAfterMinutes: parseInt(r.get('checkOutAfterMinutes') || '120'),
      crossesMidnight: String(r.get('crossesMidnight')).toLowerCase() === 'true',
      isActive: String(r.get('isActive')).toLowerCase() === 'true',
      unit: r.get('unit') || ''
    }));

    const inRecords = attendance.filter(a => a.type === 'in');
    const outRecords = attendance.filter(a => a.type === 'out');
    const now = new Date();
    const missingOuts: any[] = [];

    for (const inRec of inRecords) {
      if (outRecords.some(o => o.nip === inRec.nip && o.date === inRec.date)) continue;

      let targetShift = shifts.find(s => s.name === inRec.shift);
      if (!targetShift) {
        const empUnit = employees.find(e => e.nip === inRec.nip)?.unit || '';
        const activeShifts = shifts.filter(s => s.isActive);
        const specificShifts = activeShifts.filter(s => s.unit && s.unit === empUnit);
        targetShift = specificShifts[0] || activeShifts.filter(s => !s.unit || s.unit === 'none' || s.unit === '')[0] || shifts[0];
      }

      if (!targetShift) continue;

      const recDateObj = new Date(inRec.date);
      const isFriday = recDateObj.getDay() === 5;
      const isSaturday = recDateObj.getDay() === 6;
      let endTimeStr = targetShift.endTime;
      if (isFriday && targetShift.fridayEndTime) endTimeStr = targetShift.fridayEndTime;
      if (isSaturday && targetShift.saturdayEndTime) endTimeStr = targetShift.saturdayEndTime;

      if (!endTimeStr) continue;

      const [endHr, endMin] = endTimeStr.split(':').map(Number);
      const endDateTime = new Date(inRec.date);
      endDateTime.setHours(endHr, endMin, 0, 0);

      if (targetShift.crossesMidnight) {
        endDateTime.setDate(endDateTime.getDate() + 1);
      }

      endDateTime.setMinutes(endDateTime.getMinutes() + targetShift.checkOutAfterMinutes);

      if (now > endDateTime) {
        const autoCheckoutHour = (endHr - 1 + 24) % 24;
        const autoCheckoutTimeStr = `${autoCheckoutHour.toString().padStart(2, '0')}:${endMin.toString().padStart(2, '0')}`;

        missingOuts.push({
          id: (Date.now() + missingOuts.length).toString(),
          nip: inRec.nip,
          name: inRec.name,
          date: inRec.date,
          time: autoCheckoutTimeStr,
          type: 'out',
          location: inRec.location,
          status: 'Hadir (Pulang Cepat)',
          photoUrl: '',
          shift: targetShift.name
        });
      }
    }

    if (missingOuts.length > 0) {
      await attSheet.addRows(missingOuts);
      delete cache['attendance'];
      return res.json({ success: true, fixedCount: missingOuts.length });
    }

    return res.json({ success: true, fixedCount: 0 });
  } catch (error) {
    console.error('Error auto-checkout:', error);
    return res.status(500).json({ success: false, message: 'Gagal auto-checkout' });
  }
});

apiRouter.put('/attendance/:id', async (req: Request, res: Response) => {
  const { id } = req.params;
  const { status } = req.body;
  const doc = await getGoogleDoc();

  if (doc) {
    try {
      const sheet = await getSheet('Attendance');
      if (sheet) {
        const rows = await sheet.getRows();
        const rowToUpdate = rows.find(r => r.get('id') === id);
        if (rowToUpdate) {
          rowToUpdate.set('status', status);
          await rowToUpdate.save();
          delete cache['attendance'];
          return res.json({ success: true, message: 'Status berhasil diperbarui' });
        } else {
          return res.status(404).json({ success: false, message: 'Absensi tidak ditemukan' });
        }
      }
    } catch (error) {
      console.error('Error updating attendance status:', error);
      return res.status(500).json({ success: false, message: 'Terjadi kesalahan' });
    }
  }

  const record = db.attendance.find((a: any) => a.id === id);
  if (record) {
    record.status = status;
    persistFallbackDB();
    return res.json({ success: true, message: 'Status berhasil diperbarui' });
  }
  return res.status(404).json({ success: false, message: 'Absensi tidak ditemukan' });
});

// --- Password Reset & Forgot Password ---
apiRouter.post('/forgot-password', async (req: Request, res: Response) => {
  const { email } = req.body;
  let foundUser: any = null;
  let userType = '';

  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const adminSheet = await getSheet('Admins');
      if (adminSheet) {
        const rows = await adminSheet.getRows();
        const admin = rows.find(r => r.get('email') === email);
        if (admin) {
          foundUser = { id: admin.get('id'), name: admin.get('name'), email: admin.get('email') };
          userType = 'admin';
        }
      }

      if (!foundUser) {
        const empSheet = await getSheet('Employees');
        if (empSheet) {
          const rows = await empSheet.getRows();
          const emp = rows.find(r => r.get('email') === email);
          if (emp) {
            foundUser = { id: emp.get('id'), name: emp.get('name'), email: emp.get('email') };
            userType = 'employee';
          }
        }
      }

      if (!foundUser) {
        const userSheet = await getSheet('Users');
        if (userSheet) {
          const rows = await userSheet.getRows();
          const user = rows.find(r => r.get('email') === email);
          if (user) {
            foundUser = { id: user.get('id'), name: user.get('name'), email: user.get('email') };
            userType = 'user';
          }
        }
      }
    } catch (error) {
      console.error('Error checking email in spreadsheet:', error);
    }
  }

  if (!foundUser) {
    const user = db.users.find(u => u.email === email);
    if (user) {
      foundUser = user;
      userType = 'user';
    } else {
      const emp = db.employees.find(e => e.email === email);
      if (emp) {
        foundUser = emp;
        userType = 'employee';
      }
    }
  }

  if (!foundUser) {
    // Return success to prevent email enumeration
    return res.json({ success: true, message: 'Jika email terdaftar, tautan reset telah dikirim.' });
  }

  const token = Buffer.from(`${foundUser.id}:${userType}:${Date.now()}`).toString('base64');

  if (doc) {
    try {
      const resetSheet = await getOrCreateSheet('PasswordResets', ['token', 'userId', 'userType', 'expiresAt']);
      if (resetSheet) {
        await resetSheet.addRow({
          token,
          userId: foundUser.id,
          userType,
          expiresAt: Date.now() + 3600000 // 1 hour
        });
      }
    } catch (error) {
      console.error('Error saving reset token:', error);
    }
  }

  // Determine base application URL (respecting Vercel host or APP_URL env)
  const host = req.headers['x-forwarded-host'] || req.headers.host;
  const protocol = req.headers['x-forwarded-proto'] || req.protocol || 'https';
  const inferredBaseUrl = host ? `${protocol}://${host}` : 'http://localhost:3000';
  const baseUrl = process.env.APP_URL || inferredBaseUrl;
  const resetLink = `${baseUrl}/reset-password?token=${token}`;

  const resend = getResendClient();
  if (resend) {
    try {
      await resend.emails.send({
        from: 'Absensi Digital <noreplay@siabonmegilan.qzz.io>',
        to: foundUser.email,
        subject: 'Reset Password - Absensi Digital',
        html: `<p>Halo ${foundUser.name},</p><p>Klik tautan berikut untuk mereset password Anda:</p><p><a href="${resetLink}">${resetLink}</a></p><p>Tautan ini akan kedaluwarsa dalam 1 jam.</p>`,
      });
    } catch (error) {
      console.error('Error sending email via Resend:', error);
      return res.status(500).json({ success: false, message: 'Gagal mengirim email. Pastikan API Key Resend valid.' });
    }
  } else {
    console.log(`[MOCK EMAIL] To: ${foundUser.email}, Subject: Reset Password, Link: ${resetLink}`);
    return res.json({ success: true, message: 'Email reset password telah dikirim (Mock Mode)', mockLink: resetLink });
  }

  res.json({ success: true, message: 'Email reset password telah dikirim' });
});

apiRouter.post('/reset-password', async (req: Request, res: Response) => {
  const { token, newPassword } = req.body;
  if (!token || !newPassword) {
    return res.status(400).json({ success: false, message: 'Token dan password baru wajib diisi' });
  }

  let userId = '';
  let userType = '';
  let isValidToken = false;
  const doc = await getGoogleDoc();

  if (doc) {
    try {
      const resetSheet = await getSheet('PasswordResets');
      if (resetSheet) {
        const rows = await resetSheet.getRows();
        const tokenRow = rows.find(r => r.get('token') === token);
        if (tokenRow) {
          const expiresAt = parseInt(tokenRow.get('expiresAt'));
          if (Date.now() > expiresAt) {
            return res.status(400).json({ success: false, message: 'Token reset password telah kedaluwarsa' });
          }
          userId = tokenRow.get('userId');
          userType = tokenRow.get('userType');
          isValidToken = true;
          await tokenRow.delete();
        }
      }
    } catch (error) {
      console.error('Error verifying token:', error);
    }
  }

  if (!isValidToken) {
    try {
      const decoded = Buffer.from(token, 'base64').toString('utf-8');
      const parts = decoded.split(':');
      if (parts.length === 3) {
        userId = parts[0];
        userType = parts[1];
        const timestamp = parseInt(parts[2]);
        if (Date.now() - timestamp < 3600000) {
          isValidToken = true;
        } else {
          return res.status(400).json({ success: false, message: 'Token reset password telah kedaluwarsa' });
        }
      }
    } catch (e) {
      return res.status(400).json({ success: false, message: 'Token tidak valid' });
    }
  }

  if (!isValidToken) {
    return res.status(400).json({ success: false, message: 'Token tidak valid' });
  }

  let passwordUpdated = false;
  if (doc) {
    try {
      let sheetName = '';
      if (userType === 'admin') sheetName = 'Admins';
      else if (userType === 'employee') sheetName = 'Employees';
      else if (userType === 'user') sheetName = 'Users';

      if (sheetName) {
        const sheet = await getSheet(sheetName);
        if (sheet) {
          const rows = await sheet.getRows();
          const userRow = rows.find(r => r.get('id') === userId);
          if (userRow) {
            userRow.set('password', newPassword);
            await userRow.save();
            passwordUpdated = true;

            if (userType === 'user') {
              const userNip = userRow.get('nip');
              if (userNip) {
                const empSheet = await getSheet('Employees');
                if (empSheet) {
                  const empRows = await empSheet.getRows();
                  const empRow = empRows.find(r => r.get('nip') === userNip);
                  if (empRow) {
                    empRow.set('password', newPassword);
                    await empRow.save();
                  }
                }
              }
            }

            if (sheetName === 'Admins') delete cache['admins'];
            if (sheetName === 'Employees') delete cache['employees'];
          }
        }
      }
    } catch (error) {
      console.error('Error updating password:', error);
      return res.status(500).json({ success: false, message: 'Gagal memperbarui password' });
    }
  }

  if (!passwordUpdated) {
    if (userType === 'admin' || userType === 'user') {
      const user = db.users.find(u => u.id.toString() === userId);
      if (user) {
        user.password = newPassword;
        passwordUpdated = true;
      }
    } else if (userType === 'employee') {
      const emp = db.employees.find(e => e.id === userId);
      if (emp) {
        (emp as any).password = newPassword;
        passwordUpdated = true;
      }
    }
    persistFallbackDB();
  }

  if (passwordUpdated) {
    res.json({ success: true, message: 'Password berhasil diperbarui' });
  } else {
    res.status(404).json({ success: false, message: 'Pengguna tidak ditemukan' });
  }
});

// --- Users API ---
apiRouter.get('/users', (_req: Request, res: Response) => {
  res.json(db.users.map(u => ({ id: u.id, nip: u.nip, name: u.name, role: u.role })));
});

// --- Locations API ---
apiRouter.get('/locations', async (_req: Request, res: Response) => {
  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const locations = await getCachedData('locations', async () => {
        const sheet = await getOrCreateSheet('Locations', ['id', 'name', 'desa', 'kecamatan', 'kabupaten', 'coordinates', 'radius']);
        if (sheet) {
          const rows = await sheet.getRows();
          return rows.map(row => ({
            id: row.get('id'),
            desa: row.get('desa') || row.get('name') || '',
            kecamatan: row.get('kecamatan') || '',
            kabupaten: row.get('kabupaten') || '',
            coordinates: row.get('coordinates'),
            radius: row.get('radius') || 250
          }));
        }
        return [];
      });
      return res.json(locations);
    } catch (error) {
      console.error('Error fetching locations from spreadsheet:', error);
    }
  }
  res.json(db.locations);
});

apiRouter.post('/locations', async (req: Request, res: Response) => {
  const location = req.body;
  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const sheet = await getOrCreateSheet('Locations', ['id', 'name', 'desa', 'kecamatan', 'kabupaten', 'coordinates', 'radius']);
      if (sheet) {
        await sheet.addRow({
          id: Date.now().toString(),
          name: location.desa || '',
          desa: location.desa || '',
          kecamatan: location.kecamatan || '',
          kabupaten: location.kabupaten || '',
          coordinates: location.coordinates || '',
          radius: location.radius || 250
        });
        delete cache['locations'];
        return res.json({ success: true, message: 'Lokasi berhasil ditambahkan' });
      }
    } catch (error) {
      console.error('Error saving location to spreadsheet:', error);
    }
  }

  db.locations.push({
    id: Date.now().toString(),
    name: location.desa || '',
    desa: location.desa || '',
    kecamatan: location.kecamatan || '',
    kabupaten: location.kabupaten || '',
    coordinates: location.coordinates || '',
    radius: location.radius || 250
  });
  persistFallbackDB();
  res.json({ success: true, message: 'Lokasi berhasil ditambahkan' });
});

apiRouter.delete('/locations/:id', async (req: Request, res: Response) => {
  const { id } = req.params;
  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const sheet = await getSheet('Locations');
      if (sheet) {
        const rows = await sheet.getRows();
        const rowToDelete = rows.find(r => String(r.get('id')) === String(id));
        if (rowToDelete) {
          await rowToDelete.delete();
          delete cache['locations'];
          return res.json({ success: true, message: 'Lokasi berhasil dihapus' });
        }
      }
    } catch (error) {
      console.error('Error deleting location from spreadsheet:', error);
    }
  }

  db.locations = db.locations.filter(l => String(l.id) !== String(id));
  persistFallbackDB();
  res.json({ success: true, message: 'Lokasi berhasil dihapus' });
});

// --- Units API ---
apiRouter.get('/units', async (_req: Request, res: Response) => {
  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const units = await getCachedData('units', async () => {
        const sheet = await getOrCreateSheet('Units', ['id', 'name']);
        if (sheet) {
          const rows = await sheet.getRows();
          return rows.map(row => ({
            id: row.get('id'),
            name: row.get('name')
          }));
        }
        return [];
      });
      return res.json(units);
    } catch (error) {
      console.error('Error fetching units:', error);
    }
  }
  res.json(db.units || []);
});

apiRouter.post('/units', async (req: Request, res: Response) => {
  const unit = req.body;
  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const sheet = await getOrCreateSheet('Units', ['id', 'name']);
      if (sheet) {
        if (unit.id) {
          const rows = await sheet.getRows();
          const existingRow = rows.find(r => r.get('id') === unit.id);
          if (existingRow) {
            existingRow.set('name', unit.name || '');
            await existingRow.save();
          } else {
            await sheet.addRow(unit);
          }
        } else {
          unit.id = Date.now().toString();
          await sheet.addRow(unit);
        }
        delete cache['units'];
        return res.json({ success: true, message: 'Unit berhasil disimpan' });
      }
    } catch (error) {
      console.error('Error saving unit:', error);
    }
  }

  unit.id = unit.id || Date.now().toString();
  if (!db.units) db.units = [];
  const index = db.units.findIndex((u: any) => u.id === unit.id);
  if (index >= 0) db.units[index] = unit;
  else db.units.push(unit);
  persistFallbackDB();
  res.json({ success: true, message: 'Unit berhasil disimpan' });
});

apiRouter.delete('/units/:id', async (req: Request, res: Response) => {
  const { id } = req.params;
  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const sheet = await getSheet('Units');
      if (sheet) {
        const rows = await sheet.getRows();
        const rowToDelete = rows.find(r => String(r.get('id')) === String(id));
        if (rowToDelete) {
          await rowToDelete.delete();
          delete cache['units'];
          return res.json({ success: true, message: 'Unit berhasil dihapus' });
        }
      }
    } catch (error) {
      console.error('Error deleting unit:', error);
    }
  }

  if (db.units) {
    db.units = db.units.filter((u: any) => String(u.id) !== String(id));
    persistFallbackDB();
  }
  res.json({ success: true, message: 'Unit berhasil dihapus' });
});

// --- Shifts API ---
apiRouter.get('/shifts', async (_req: Request, res: Response) => {
  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const shifts = await getCachedData('shifts', async () => {
        const sheet = await getOrCreateSheet('Shifts', ['id', 'name', 'startTime', 'endTime', 'fridayEndTime', 'saturdayEndTime', 'checkInBeforeMinutes', 'checkInAfterMinutes', 'checkOutBeforeMinutes', 'checkOutAfterMinutes', 'crossesMidnight', 'isActive', 'unit', 'isOffSunday', 'isOffHoliday']);
        if (sheet) {
          const rows = await sheet.getRows();
          return rows.map(row => ({
            id: row.get('id'),
            name: row.get('name'),
            startTime: row.get('startTime'),
            endTime: row.get('endTime'),
            fridayEndTime: row.get('fridayEndTime') || '',
            saturdayEndTime: row.get('saturdayEndTime') || '',
            checkInBeforeMinutes: parseInt(row.get('checkInBeforeMinutes') || '60'),
            checkInAfterMinutes: parseInt(row.get('checkInAfterMinutes') || '15'),
            checkOutBeforeMinutes: parseInt(row.get('checkOutBeforeMinutes') || '10'),
            checkOutAfterMinutes: parseInt(row.get('checkOutAfterMinutes') || '120'),
            crossesMidnight: String(row.get('crossesMidnight')).toLowerCase() === 'true',
            isActive: String(row.get('isActive')).toLowerCase() === 'true',
            unit: row.get('unit') || '',
            isOffSunday: String(row.get('isOffSunday')).toLowerCase() === 'true',
            isOffHoliday: String(row.get('isOffHoliday')).toLowerCase() === 'true'
          }));
        }
        return [];
      });
      return res.json(shifts);
    } catch (error) {
      console.error('Error fetching shifts from spreadsheet:', error);
    }
  }
  res.json(db.shifts || []);
});

apiRouter.post('/shifts', async (req: Request, res: Response) => {
  const shift = req.body;
  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const sheet = await getOrCreateSheet('Shifts', ['id', 'name', 'startTime', 'endTime', 'fridayEndTime', 'saturdayEndTime', 'checkInBeforeMinutes', 'checkInAfterMinutes', 'checkOutBeforeMinutes', 'checkOutAfterMinutes', 'crossesMidnight', 'isActive', 'unit', 'isOffSunday', 'isOffHoliday']);
      if (sheet) {
        await sheet.addRow({
          ...shift,
          checkInBeforeMinutes: (shift.checkInBeforeMinutes || 60).toString(),
          checkInAfterMinutes: (shift.checkInAfterMinutes || 15).toString(),
          checkOutBeforeMinutes: (shift.checkOutBeforeMinutes || 10).toString(),
          checkOutAfterMinutes: (shift.checkOutAfterMinutes || 120).toString(),
          crossesMidnight: (shift.crossesMidnight || false).toString(),
          isActive: (shift.isActive ?? true).toString(),
          unit: shift.unit || '',
          isOffSunday: (shift.isOffSunday || false).toString(),
          isOffHoliday: (shift.isOffHoliday || false).toString()
        });
        delete cache['shifts'];
        return res.json({ success: true, message: 'Shift berhasil ditambahkan' });
      }
    } catch (error) {
      console.error('Error saving shift to spreadsheet:', error);
    }
  }

  if (!db.shifts) db.shifts = [];
  db.shifts.push({ id: Date.now().toString(), ...shift });
  persistFallbackDB();
  res.json({ success: true, message: 'Shift berhasil ditambahkan' });
});

apiRouter.delete('/shifts/:id', async (req: Request, res: Response) => {
  const { id } = req.params;
  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const sheet = await getSheet('Shifts');
      if (sheet) {
        const rows = await sheet.getRows();
        const rowToDelete = rows.find(r => r.get('id') === id);
        if (rowToDelete) {
          await rowToDelete.delete();
          delete cache['shifts'];
          return res.json({ success: true, message: 'Shift berhasil dihapus' });
        }
      }
    } catch (error) {
      console.error('Error deleting shift from spreadsheet:', error);
    }
  }

  if (db.shifts) {
    db.shifts = db.shifts.filter(s => String(s.id) !== String(id));
    persistFallbackDB();
  }
  res.json({ success: true, message: 'Shift berhasil dihapus' });
});

apiRouter.put('/shifts/:id', async (req: Request, res: Response) => {
  const { id } = req.params;
  const shift = req.body;
  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const sheet = await getSheet('Shifts');
      if (sheet) {
        const rows = await sheet.getRows();
        const rowToUpdate = rows.find(r => String(r.get('id')) === String(id));
        if (rowToUpdate) {
          rowToUpdate.set('name', shift.name);
          rowToUpdate.set('startTime', shift.startTime);
          rowToUpdate.set('endTime', shift.endTime);
          rowToUpdate.set('fridayEndTime', shift.fridayEndTime || '');
          rowToUpdate.set('saturdayEndTime', shift.saturdayEndTime || '');
          rowToUpdate.set('checkInBeforeMinutes', (shift.checkInBeforeMinutes || 60).toString());
          rowToUpdate.set('checkInAfterMinutes', (shift.checkInAfterMinutes || 15).toString());
          rowToUpdate.set('checkOutBeforeMinutes', (shift.checkOutBeforeMinutes || 10).toString());
          rowToUpdate.set('checkOutAfterMinutes', (shift.checkOutAfterMinutes || 120).toString());
          rowToUpdate.set('crossesMidnight', (shift.crossesMidnight || false).toString());
          rowToUpdate.set('isActive', (shift.isActive ?? true).toString());
          rowToUpdate.set('unit', shift.unit || '');
          rowToUpdate.set('isOffSunday', (shift.isOffSunday || false).toString());
          rowToUpdate.set('isOffHoliday', (shift.isOffHoliday || false).toString());
          await rowToUpdate.save();
          delete cache['shifts'];
          return res.json({ success: true, message: 'Shift berhasil diperbarui' });
        }
      }
    } catch (error) {
      console.error('Error updating shift in spreadsheet:', error);
      return res.status(500).json({ success: false, message: 'Gagal memperbarui shift di spreadsheet' });
    }
  }

  if (db.shifts) {
    const idx = db.shifts.findIndex(s => String(s.id) === String(id));
    if (idx >= 0) {
      db.shifts[idx] = { ...db.shifts[idx], ...shift };
      persistFallbackDB();
    }
  }
  res.json({ success: true, message: 'Shift berhasil diperbarui' });
});

// --- Announcements API ---
apiRouter.get('/announcements', async (_req: Request, res: Response) => {
  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const announcements = await getCachedData('announcements', async () => {
        const sheet = await getOrCreateSheet('Announcements', ['id', 'title', 'content', 'date', 'expiryDate', 'isActive']);
        if (sheet) {
          const rows = await sheet.getRows();
          return rows.map(row => ({
            id: row.get('id'),
            title: row.get('title'),
            content: row.get('content'),
            date: row.get('date'),
            expiryDate: row.get('expiryDate'),
            isActive: String(row.get('isActive')).toLowerCase() === 'true'
          }));
        }
        return [];
      });
      return res.json(announcements);
    } catch (error) {
      console.error('Error fetching announcements from spreadsheet:', error);
    }
  }
  res.json(db.announcements || []);
});

apiRouter.post('/announcements', async (req: Request, res: Response) => {
  const announcement = req.body;
  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const sheet = await getOrCreateSheet('Announcements', ['id', 'title', 'content', 'date', 'expiryDate', 'isActive']);
      if (sheet) {
        await sheet.addRow({
          ...announcement,
          isActive: (announcement.isActive ?? true).toString(),
          date: announcement.date || new Date().toISOString().split('T')[0]
        });
        delete cache['announcements'];
        return res.json({ success: true, message: 'Pengumuman berhasil ditambahkan' });
      }
    } catch (error) {
      console.error('Error saving announcement to spreadsheet:', error);
      return res.status(500).json({ success: false, message: 'Gagal menyimpan pengumuman' });
    }
  }

  if (!db.announcements) db.announcements = [];
  db.announcements.push({ id: Date.now().toString(), ...announcement });
  persistFallbackDB();
  res.json({ success: true, message: 'Pengumuman berhasil ditambahkan' });
});

apiRouter.put('/announcements/:id', async (req: Request, res: Response) => {
  const { id } = req.params;
  const { isActive } = req.body;
  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const sheet = await getSheet('Announcements');
      if (sheet) {
        const rows = await sheet.getRows();
        const rowToUpdate = rows.find(r => r.get('id') === id);
        if (rowToUpdate) {
          rowToUpdate.set('isActive', (isActive ?? true).toString());
          await rowToUpdate.save();
          delete cache['announcements'];
          return res.json({ success: true, message: 'Pengumuman diupdate' });
        }
      }
    } catch (error) {
      console.error('Error updating announcement:', error);
      return res.status(500).json({ success: false, message: 'Gagal update pengumuman' });
    }
  }

  if (db.announcements) {
    const item = db.announcements.find(a => String(a.id) === String(id));
    if (item) item.isActive = isActive;
    persistFallbackDB();
  }
  res.json({ success: true, message: 'Pengumuman diupdate' });
});

apiRouter.delete('/announcements/:id', async (req: Request, res: Response) => {
  const { id } = req.params;
  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const sheet = await getSheet('Announcements');
      if (sheet) {
        const rows = await sheet.getRows();
        const rowToDelete = rows.find(r => r.get('id') === id);
        if (rowToDelete) {
          await rowToDelete.delete();
          delete cache['announcements'];
          return res.json({ success: true, message: 'Pengumuman dihapus' });
        }
      }
    } catch (error) {
      console.error('Error deleting announcement:', error);
      return res.status(500).json({ success: false, message: 'Gagal menghapus pengumuman' });
    }
  }

  if (db.announcements) {
    db.announcements = db.announcements.filter(a => String(a.id) !== String(id));
    persistFallbackDB();
  }
  res.json({ success: true, message: 'Pengumuman dihapus' });
});

// --- Holidays API ---
apiRouter.get('/holidays', async (_req: Request, res: Response) => {
  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const items = await getCachedData('holidays', async () => {
        const sheet = await getOrCreateSheet('Holidays', ['id', 'date', 'name']);
        if (sheet) {
          const rows = await sheet.getRows();
          return rows.map(row => ({
            id: row.get('id'),
            date: row.get('date'),
            name: row.get('name')
          }));
        }
        return [];
      });
      return res.json(items);
    } catch (error) {
      console.error('Error fetching holidays from spreadsheet:', error);
    }
  }
  res.json(db.holidays || []);
});

apiRouter.post('/holidays', async (req: Request, res: Response) => {
  const item = req.body;
  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const sheet = await getOrCreateSheet('Holidays', ['id', 'date', 'name']);
      if (sheet) {
        await sheet.addRow({
          id: item.id || Date.now().toString(),
          date: item.date,
          name: item.name
        });
        delete cache['holidays'];
        return res.json({ success: true, message: 'Added successfully' });
      }
    } catch (error) {
      console.error('Error adding holiday to spreadsheet:', error);
      return res.status(500).json({ success: false, message: 'Failed to add item' });
    }
  }

  if (!db.holidays) db.holidays = [];
  db.holidays.push({ id: item.id || Date.now().toString(), ...item });
  persistFallbackDB();
  res.json({ success: true, message: 'Added successfully' });
});

apiRouter.delete('/holidays/:id', async (req: Request, res: Response) => {
  const { id } = req.params;
  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const sheet = await getSheet('Holidays');
      if (sheet) {
        const rows = await sheet.getRows();
        const rowToDelete = rows.find(r => String(r.get('id')) === String(id));
        if (rowToDelete) {
          await rowToDelete.delete();
          delete cache['holidays'];
          return res.json({ success: true });
        } else {
          return res.status(404).json({ success: false });
        }
      }
    } catch (error) {
      return res.status(500).json({ success: false });
    }
  }

  if (db.holidays) {
    db.holidays = db.holidays.filter(h => String(h.id) !== String(id));
    persistFallbackDB();
  }
  res.json({ success: true });
});

// --- Settings API ---
apiRouter.get('/settings', async (_req: Request, res: Response) => {
  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const settingsPayload = await getCachedData('settings', async () => {
        const sheet = await getOrCreateSheet('Settings', ['key', 'value']);
        if (sheet) {
          const rows = await sheet.getRows();
          const settings: any = {};
          rows.forEach(row => {
            try {
              settings[row.get('key')] = JSON.parse(row.get('value'));
            } catch (e) {
              settings[row.get('key')] = row.get('value');
            }
          });
          return settings;
        }
        return {};
      });
      if (Object.keys(settingsPayload).length > 0) {
        return res.json(settingsPayload);
      }
    } catch (error) {
      console.error('Error fetching settings from spreadsheet:', error);
    }
  }
  res.json(db.settings);
});

apiRouter.post('/settings', async (req: Request, res: Response) => {
  const { key, value } = req.body;
  db.settings = { ...db.settings, [key]: value };
  persistFallbackDB();

  const doc = await getGoogleDoc();
  if (doc) {
    try {
      const sheet = await getOrCreateSheet('Settings', ['key', 'value']);
      if (sheet) {
        const rows = await sheet.getRows();
        const existingRow = rows.find(r => r.get('key') === key);
        let stringifiedValue = JSON.stringify(value);
        if (stringifiedValue.length > 45000) {
          console.warn(`Settings value for ${key} is very large (${stringifiedValue.length} chars). Truncating for Google Sheets cell safety.`);
          stringifiedValue = stringifiedValue.substring(0, 45000);
        }

        if (existingRow) {
          existingRow.set('value', stringifiedValue);
          await existingRow.save();
        } else {
          await sheet.addRow({ key, value: stringifiedValue });
        }
        delete cache['settings'];
      }
    } catch (error) {
      console.error('Error saving settings to spreadsheet:', error);
      return res.status(500).json({ success: false, message: 'Gagal menyimpan ke spreadsheet. Mungkin ukuran data terlalu besar (misal: gambar logo).' });
    }
  }
  res.json({ success: true, message: 'Pengaturan berhasil disimpan' });
});

// Mount the API Router on both '/api' and '/' for bulletproof Vercel routing
app.use('/api', apiRouter);
app.use('/', apiRouter);

// Global Error Handler
app.use((err: any, _req: Request, res: Response, _next: NextFunction) => {
  console.error('Unhandled Application Error:', err);
  res.status(500).json({
    success: false,
    message: 'Terjadi kesalahan internal pada server.',
    error: process.env.NODE_ENV === 'development' ? err?.message : undefined
  });
});

export { app, db };
export default app;
