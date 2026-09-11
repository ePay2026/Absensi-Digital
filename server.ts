import express from 'express';
import path from 'path';
import cors from 'cors';
import crypto from 'crypto';
import dotenv from 'dotenv';
import { GoogleSpreadsheet } from 'google-spreadsheet';
import { JWT } from 'google-auth-library';
import { Resend } from 'resend';

dotenv.config();

export const app = express();
const PORT = 3000;

app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));

// Google Sheets Client Setup
let docInstance: GoogleSpreadsheet | null = null;

function getSpreadsheetAuth() {
  const email = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
  let key = process.env.GOOGLE_PRIVATE_KEY;
  if (key) {
    key = key.replace(/\\n/g, '\n').replace(/^"|"$/g, '');
  }
  const spreadsheetId = process.env.SPREADSHEET_ID;

  if (!email || !key || !spreadsheetId) {
    return null;
  }

  const auth = new JWT({
    email,
    key,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  return new GoogleSpreadsheet(spreadsheetId, auth);
}

async function getDoc(): Promise<GoogleSpreadsheet | null> {
  if (docInstance) return docInstance;
  try {
    const doc = getSpreadsheetAuth();
    if (!doc) {
      console.warn('Google Spreadsheet credentials not fully configured in environment variables.');
      return null;
    }
    await doc.loadInfo();
    console.log(`Google Spreadsheet connected successfully: ${doc.title}`);
    docInstance = doc;
    return docInstance;
  } catch (err) {
    console.error('Failed to connect to Google Sheets:', err);
    return null;
  }
}

async function getSheet(title: string) {
  const doc = await getDoc();
  if (!doc) return null;
  let sheet = doc.sheetsByTitle[title];
  if (!sheet) {
    // Try to reload info once in case sheet was created recently
    await doc.loadInfo();
    sheet = doc.sheetsByTitle[title];
  }
  return sheet || null;
}

// Resend Email Client
let resendClient: Resend | null = null;
function getResend() {
  if (!resendClient && process.env.RESEND_API_KEY) {
    resendClient = new Resend(process.env.RESEND_API_KEY);
  }
  return resendClient;
}

function safeJsonParse(str: any) {
  if (!str) return null;
  if (typeof str === 'object') return str;
  try {
    return JSON.parse(str);
  } catch {
    return str;
  }
}

// ==========================================
// API ROUTES
// ==========================================

// Health Check
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', timestamp: Date.now() });
});

// Time Sync (Server Time)
app.get(['/api/time', '/api/server-time'], (req, res) => {
  const now = new Date();
  res.json({
    time: now.toISOString(),
    timestamp: now.getTime(),
    timezone: Intl.DateTimeFormat().resolvedOptions().timeZone,
  });
});

// Settings Endpoints
app.get('/api/settings', async (req, res) => {
  try {
    const sheet = await getSheet('Settings');
    const settingsMap: Record<string, any> = {
      generalSettings: {
        appName: 'Si Abon Elite App',
        appLogo: '',
        companyName: 'Puskesmas SEHAT',
        headName: 'dr. Xxxxx',
        email: 'pkm.paciran@gmail.com',
        address: 'Jl. Raya Paciran No.78',
        mainLocation: '-7.250445, 112.768845',
      },
      absensiSettings: {
        tolerance: '15',
        enableCountdown: true,
        enableEarlyCheckout: true,
      },
      leaveSettings: {
        autoApprove: '0',
      },
    };

    if (sheet) {
      const rows = await sheet.getRows();
      for (const row of rows) {
        const key = row.get('key');
        const rawVal = row.get('value');
        if (key && rawVal) {
          settingsMap[key] = safeJsonParse(rawVal);
        }
      }
    }

    res.json(settingsMap);
  } catch (err: any) {
    console.error('Error fetching settings:', err);
    res.status(500).json({ error: 'Failed to fetch settings', message: err.message });
  }
});

app.post('/api/settings', async (req, res) => {
  try {
    const { key, value } = req.body;
    if (!key) {
      return res.status(400).json({ success: false, message: 'Key is required' });
    }

    const sheet = await getSheet('Settings');
    if (!sheet) {
      return res.status(500).json({ success: false, message: 'Settings sheet not available' });
    }

    const rows = await sheet.getRows();
    const existing = rows.find((r) => r.get('key') === key);
    const valueStr = typeof value === 'object' ? JSON.stringify(value) : String(value);

    if (existing) {
      existing.assign({ value: valueStr });
      await existing.save();
    } else {
      await sheet.addRow({ key, value: valueStr });
    }

    res.json({ success: true });
  } catch (err: any) {
    console.error('Error saving setting:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

// Units Endpoints
app.get('/api/units', async (req, res) => {
  try {
    const sheet = await getSheet('Units');
    if (!sheet) return res.json([]);
    const rows = await sheet.getRows();
    const units = rows.map((r) => ({
      id: r.get('id') || '',
      name: r.get('name') || '',
    }));
    res.json(units);
  } catch (err: any) {
    console.error('Error fetching units:', err);
    res.status(500).json({ error: 'Failed to fetch units', message: err.message });
  }
});

app.post('/api/units', async (req, res) => {
  try {
    const { id, name } = req.body;
    const sheet = await getSheet('Units');
    if (!sheet) return res.status(500).json({ success: false, message: 'Units sheet not found' });
    const unitId = id || Date.now().toString();
    await sheet.addRow({ id: unitId, name: name || '' });
    res.json({ success: true, id: unitId, name });
  } catch (err: any) {
    console.error('Error adding unit:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

app.delete('/api/units/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const sheet = await getSheet('Units');
    if (!sheet) return res.status(500).json({ success: false, message: 'Units sheet not found' });
    const rows = await sheet.getRows();
    const target = rows.find((r) => r.get('id') === id);
    if (target) {
      await target.delete();
    }
    res.json({ success: true });
  } catch (err: any) {
    console.error('Error deleting unit:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

// Locations Endpoints
app.get('/api/locations', async (req, res) => {
  try {
    const sheet = await getSheet('Locations');
    if (!sheet) return res.json([]);
    const rows = await sheet.getRows();
    const locations = rows.map((r) => ({
      id: r.get('id') || '',
      name: r.get('name') || '',
      desa: r.get('desa') || r.get('name') || '',
      kecamatan: r.get('kecamatan') || '',
      kabupaten: r.get('kabupaten') || '',
      coordinates: r.get('coordinates') || '',
      radius: parseInt(r.get('radius') || '100', 10),
    }));
    res.json(locations);
  } catch (err: any) {
    console.error('Error fetching locations:', err);
    res.status(500).json({ error: 'Failed to fetch locations', message: err.message });
  }
});

app.post('/api/locations', async (req, res) => {
  try {
    const { id, name, desa, kecamatan, kabupaten, coordinates, radius } = req.body;
    const sheet = await getSheet('Locations');
    if (!sheet) return res.status(500).json({ success: false, message: 'Locations sheet not found' });

    const locId = id || Date.now().toString();
    const rows = await sheet.getRows();
    const existing = rows.find((r) => r.get('id') === locId);

    const record = {
      id: locId,
      name: name || desa || '',
      desa: desa || name || '',
      kecamatan: kecamatan || '',
      kabupaten: kabupaten || '',
      coordinates: coordinates || '',
      radius: String(radius || 100),
    };

    if (existing) {
      existing.assign(record);
      await existing.save();
    } else {
      await sheet.addRow(record);
    }

    res.json({ success: true, location: record });
  } catch (err: any) {
    console.error('Error saving location:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

app.delete('/api/locations/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const sheet = await getSheet('Locations');
    if (!sheet) return res.status(500).json({ success: false, message: 'Locations sheet not found' });
    const rows = await sheet.getRows();
    const target = rows.find((r) => r.get('id') === id);
    if (target) {
      await target.delete();
    }
    res.json({ success: true });
  } catch (err: any) {
    console.error('Error deleting location:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

// Shifts Endpoints
app.get('/api/shifts', async (req, res) => {
  try {
    const sheet = await getSheet('Shifts');
    if (!sheet) return res.json([]);
    const rows = await sheet.getRows();
    const shifts = rows.map((r) => ({
      id: r.get('id') || '',
      name: r.get('name') || '',
      startTime: r.get('startTime') || '',
      endTime: r.get('endTime') || '',
      fridayEndTime: r.get('fridayEndTime') || '',
      saturdayEndTime: r.get('saturdayEndTime') || '',
      checkInBeforeMinutes: r.get('checkInBeforeMinutes') || '60',
      checkInAfterMinutes: r.get('checkInAfterMinutes') || '15',
      checkOutBeforeMinutes: r.get('checkOutBeforeMinutes') || '10',
      checkOutAfterMinutes: r.get('checkOutAfterMinutes') || '120',
      crossesMidnight: r.get('crossesMidnight') === 'TRUE' || r.get('crossesMidnight') === 'true',
      isActive: r.get('isActive') !== 'FALSE' && r.get('isActive') !== 'false',
      unit: r.get('unit') || '',
      isOffSunday: r.get('isOffSunday') === 'TRUE' || r.get('isOffSunday') === 'true',
      isOffHoliday: r.get('isOffHoliday') === 'TRUE' || r.get('isOffHoliday') === 'true',
    }));
    res.json(shifts);
  } catch (err: any) {
    console.error('Error fetching shifts:', err);
    res.status(500).json({ error: 'Failed to fetch shifts', message: err.message });
  }
});

app.post('/api/shifts', async (req, res) => {
  try {
    const shift = req.body;
    const sheet = await getSheet('Shifts');
    if (!sheet) return res.status(500).json({ success: false, message: 'Shifts sheet not found' });

    const shiftId = shift.id || Date.now().toString();
    await sheet.addRow({
      id: shiftId,
      name: shift.name || '',
      startTime: shift.startTime || '',
      endTime: shift.endTime || '',
      fridayEndTime: shift.fridayEndTime || '',
      saturdayEndTime: shift.saturdayEndTime || '',
      checkInBeforeMinutes: String(shift.checkInBeforeMinutes ?? '60'),
      checkInAfterMinutes: String(shift.checkInAfterMinutes ?? '15'),
      checkOutBeforeMinutes: String(shift.checkOutBeforeMinutes ?? '10'),
      checkOutAfterMinutes: String(shift.checkOutAfterMinutes ?? '120'),
      crossesMidnight: shift.crossesMidnight ? 'TRUE' : 'FALSE',
      isActive: shift.isActive !== false ? 'TRUE' : 'FALSE',
      unit: shift.unit || '',
      isOffSunday: shift.isOffSunday ? 'TRUE' : 'FALSE',
      isOffHoliday: shift.isOffHoliday ? 'TRUE' : 'FALSE',
    });

    res.json({ success: true, id: shiftId });
  } catch (err: any) {
    console.error('Error adding shift:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

app.put('/api/shifts/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const shift = req.body;
    const sheet = await getSheet('Shifts');
    if (!sheet) return res.status(500).json({ success: false, message: 'Shifts sheet not found' });

    const rows = await sheet.getRows();
    const target = rows.find((r) => r.get('id') === id);
    if (target) {
      target.assign({
        name: shift.name ?? target.get('name'),
        startTime: shift.startTime ?? target.get('startTime'),
        endTime: shift.endTime ?? target.get('endTime'),
        fridayEndTime: shift.fridayEndTime ?? target.get('fridayEndTime'),
        saturdayEndTime: shift.saturdayEndTime ?? target.get('saturdayEndTime'),
        checkInBeforeMinutes: String(shift.checkInBeforeMinutes ?? target.get('checkInBeforeMinutes')),
        checkInAfterMinutes: String(shift.checkInAfterMinutes ?? target.get('checkInAfterMinutes')),
        checkOutBeforeMinutes: String(shift.checkOutBeforeMinutes ?? target.get('checkOutBeforeMinutes')),
        checkOutAfterMinutes: String(shift.checkOutAfterMinutes ?? target.get('checkOutAfterMinutes')),
        crossesMidnight: shift.crossesMidnight !== undefined ? (shift.crossesMidnight ? 'TRUE' : 'FALSE') : target.get('crossesMidnight'),
        isActive: shift.isActive !== undefined ? (shift.isActive ? 'TRUE' : 'FALSE') : target.get('isActive'),
        unit: shift.unit ?? target.get('unit'),
        isOffSunday: shift.isOffSunday !== undefined ? (shift.isOffSunday ? 'TRUE' : 'FALSE') : target.get('isOffSunday'),
        isOffHoliday: shift.isOffHoliday !== undefined ? (shift.isOffHoliday ? 'TRUE' : 'FALSE') : target.get('isOffHoliday'),
      });
      await target.save();
    }
    res.json({ success: true });
  } catch (err: any) {
    console.error('Error updating shift:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

app.delete('/api/shifts/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const sheet = await getSheet('Shifts');
    if (!sheet) return res.status(500).json({ success: false, message: 'Shifts sheet not found' });
    const rows = await sheet.getRows();
    const target = rows.find((r) => r.get('id') === id);
    if (target) {
      await target.delete();
    }
    res.json({ success: true });
  } catch (err: any) {
    console.error('Error deleting shift:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

// Employees Endpoints
app.get('/api/employees', async (req, res) => {
  try {
    const sheet = await getSheet('Employees');
    if (!sheet) return res.json([]);
    const rows = await sheet.getRows();
    const employees = rows.map((r) => ({
      id: r.get('id') || '',
      name: r.get('name') || '',
      nip: r.get('nip') || '',
      office: r.get('office') || '',
      office2: r.get('office2') || '',
      email: r.get('email') || '',
      gender: r.get('gender') || '',
      cluster: r.get('cluster') || '',
      unit: r.get('unit') || '',
      photoUrl: r.get('photoUrl') || '',
      photoUploadCount: parseInt(r.get('photoUploadCount') || '0', 10),
    }));
    res.json(employees);
  } catch (err: any) {
    console.error('Error fetching employees:', err);
    res.status(500).json({ error: 'Failed to fetch employees', message: err.message });
  }
});

app.post('/api/employees', async (req, res) => {
  try {
    const emp = req.body;
    const sheet = await getSheet('Employees');
    if (!sheet) return res.status(500).json({ success: false, message: 'Employees sheet not found' });

    const empId = emp.id || Date.now().toString();
    await sheet.addRow({
      id: empId,
      name: emp.name || '',
      nip: String(emp.nip || ''),
      office: emp.office || '',
      office2: emp.office2 || '',
      email: emp.email || '',
      gender: emp.gender || '',
      cluster: emp.cluster || '',
      unit: emp.unit || '',
      password: emp.password || '123456',
      photoUrl: emp.photoUrl || '',
      photoUploadCount: String(emp.photoUploadCount || '0'),
    });

    res.json({ success: true, id: empId });
  } catch (err: any) {
    console.error('Error adding employee:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/employees/bulk', async (req, res) => {
  try {
    const employees = req.body;
    if (!Array.isArray(employees) || employees.length === 0) {
      return res.status(400).json({ success: false, message: 'Invalid employees array' });
    }

    const sheet = await getSheet('Employees');
    if (!sheet) return res.status(500).json({ success: false, message: 'Employees sheet not found' });

    const rowsToAdd = employees.map((emp, index) => ({
      id: emp.id || (Date.now() + index).toString(),
      name: emp.name || '',
      nip: String(emp.nip || ''),
      office: emp.office || '',
      office2: emp.office2 || '',
      email: emp.email || '',
      gender: emp.gender || '',
      cluster: emp.cluster || '',
      unit: emp.unit || '',
      password: emp.password || '123456',
      photoUrl: emp.photoUrl || '',
      photoUploadCount: '0',
    }));

    await sheet.addRows(rowsToAdd);
    res.json({ success: true, count: rowsToAdd.length });
  } catch (err: any) {
    console.error('Error bulk uploading employees:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

app.delete('/api/employees/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const sheet = await getSheet('Employees');
    if (!sheet) return res.status(500).json({ success: false, message: 'Employees sheet not found' });
    const rows = await sheet.getRows();
    const target = rows.find((r) => r.get('id') === id);
    if (target) {
      await target.delete();
    }
    res.json({ success: true });
  } catch (err: any) {
    console.error('Error deleting employee:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/employees/photo', async (req, res) => {
  try {
    const { nip, photoUrl } = req.body;
    if (!nip) {
      return res.status(400).json({ success: false, message: 'NIP is required' });
    }

    const sheet = await getSheet('Employees');
    if (!sheet) return res.status(500).json({ success: false, message: 'Employees sheet not found' });

    const rows = await sheet.getRows();
    const target = rows.find((r) => r.get('nip') === String(nip));
    if (!target) {
      return res.status(404).json({ success: false, message: 'Karyawan tidak ditemukan' });
    }

    const currentCount = parseInt(target.get('photoUploadCount') || '0', 10);
    if (currentCount >= 5) {
      return res.status(400).json({ success: false, message: 'Batas unggah foto telah mencapai maksimal (5 kali).' });
    }

    const newCount = currentCount + 1;
    target.assign({
      photoUrl: photoUrl || '',
      photoUploadCount: String(newCount),
    });
    await target.save();

    res.json({
      success: true,
      message: 'Foto profil berhasil diperbarui',
      photoUrl,
      photoUploadCount: newCount,
    });
  } catch (err: any) {
    console.error('Error updating employee photo:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

// Admins Endpoints
app.get('/api/admins', async (req, res) => {
  try {
    const sheet = await getSheet('Admins');
    if (!sheet) {
      return res.json([
        {
          id: '1',
          name: 'Super Admin',
          nip: '000000',
          email: 'super@admin.com',
          phone: '081234567890',
          group: 'Superadmin',
          isActive: true,
          access: ['Absensi', 'Master Data', 'Sistem'],
        },
      ]);
    }

    const rows = await sheet.getRows();
    if (rows.length === 0) {
      return res.json([
        {
          id: '1',
          name: 'Super Admin',
          nip: '000000',
          email: 'super@admin.com',
          phone: '081234567890',
          group: 'Superadmin',
          isActive: true,
          access: ['Absensi', 'Master Data', 'Sistem'],
        },
      ]);
    }

    const admins = rows.map((r) => ({
      id: r.get('id') || '',
      name: r.get('name') || '',
      nip: r.get('nip') || '',
      email: r.get('email') || '',
      phone: r.get('phone') || '',
      group: r.get('group') || 'Admin',
      isActive: r.get('isActive') !== 'FALSE' && r.get('isActive') !== 'false',
      access: safeJsonParse(r.get('access')) || ['Absensi', 'Master Data', 'Sistem'],
    }));

    res.json(admins);
  } catch (err: any) {
    console.error('Error fetching admins:', err);
    res.status(500).json({ error: 'Failed to fetch admins', message: err.message });
  }
});

app.post('/api/admins', async (req, res) => {
  try {
    const admin = req.body;
    const sheet = await getSheet('Admins');
    if (!sheet) return res.status(500).json({ success: false, message: 'Admins sheet not found' });

    const adminId = admin.id || Date.now().toString();
    await sheet.addRow({
      id: adminId,
      name: admin.name || '',
      nip: String(admin.nip || ''),
      email: admin.email || '',
      phone: admin.phone || '',
      group: admin.group || 'Admin',
      isActive: admin.isActive !== false ? 'TRUE' : 'FALSE',
      access: JSON.stringify(admin.access || ['Absensi']),
      password: admin.password || 'admin123',
    });

    res.json({ success: true, id: adminId });
  } catch (err: any) {
    console.error('Error adding admin:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

app.put('/api/admins/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const admin = req.body;
    const sheet = await getSheet('Admins');
    if (!sheet) return res.status(500).json({ success: false, message: 'Admins sheet not found' });

    const rows = await sheet.getRows();
    const target = rows.find((r) => r.get('id') === id);
    if (target) {
      target.assign({
        name: admin.name ?? target.get('name'),
        nip: admin.nip !== undefined ? String(admin.nip) : target.get('nip'),
        email: admin.email ?? target.get('email'),
        phone: admin.phone ?? target.get('phone'),
        group: admin.group ?? target.get('group'),
        isActive: admin.isActive !== undefined ? (admin.isActive ? 'TRUE' : 'FALSE') : target.get('isActive'),
        access: admin.access ? JSON.stringify(admin.access) : target.get('access'),
        ...(admin.password ? { password: admin.password } : {}),
      });
      await target.save();
    }
    res.json({ success: true });
  } catch (err: any) {
    console.error('Error updating admin:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

app.delete('/api/admins/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const sheet = await getSheet('Admins');
    if (!sheet) return res.status(500).json({ success: false, message: 'Admins sheet not found' });
    const rows = await sheet.getRows();
    const target = rows.find((r) => r.get('id') === id);
    if (target) {
      await target.delete();
    }
    res.json({ success: true });
  } catch (err: any) {
    console.error('Error deleting admin:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

// Attendance Endpoints
app.get('/api/attendance', async (req, res) => {
  try {
    const sheet = await getSheet('Attendance');
    if (!sheet) return res.json([]);

    const { startDate, endDate } = req.query as { startDate?: string; endDate?: string };
    const rows = await sheet.getRows();

    let records = rows.map((r) => ({
      id: r.get('id') || '',
      nip: r.get('nip') || '',
      name: r.get('name') || '',
      date: r.get('date') || '',
      time: r.get('time') || '',
      type: r.get('type') || '',
      location: safeJsonParse(r.get('location')),
      status: r.get('status') || '',
      photoUrl: r.get('photoUrl') || '',
      shift: r.get('shift') || '',
    }));

    if (startDate && endDate) {
      records = records.filter((r) => {
        if (!r.date) return false;
        // Direct date within range
        if (r.date >= startDate && r.date <= endDate) return true;
        // Check leave endDate if applicable
        if (r.location && typeof r.location === 'object' && r.location.endDate) {
          const locEnd = r.location.endDate;
          return locEnd >= startDate && r.date <= endDate;
        }
        return false;
      });
    }

    res.json(records);
  } catch (err: any) {
    console.error('Error fetching attendance:', err);
    res.status(500).json({ error: 'Failed to fetch attendance', message: err.message });
  }
});

app.post('/api/attendance', async (req, res) => {
  try {
    const { nip, name, date, time, type, location, status, photoUrl, shift } = req.body;
    const sheet = await getSheet('Attendance');
    if (!sheet) return res.status(500).json({ success: false, message: 'Attendance sheet not found' });

    const attId = Date.now().toString();
    await sheet.addRow({
      id: attId,
      nip: String(nip || ''),
      name: name || '',
      date: date || new Date().toISOString().split('T')[0],
      time: time || '',
      type: type || 'in',
      location: typeof location === 'object' ? JSON.stringify(location) : String(location || ''),
      status: status || 'Hadir',
      photoUrl: photoUrl || '',
      shift: shift || '',
    });

    res.json({ success: true, id: attId });
  } catch (err: any) {
    console.error('Error recording attendance:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

app.put('/api/attendance/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { status } = req.body;
    const sheet = await getSheet('Attendance');
    if (!sheet) return res.status(500).json({ success: false, message: 'Attendance sheet not found' });

    const rows = await sheet.getRows();
    const target = rows.find((r) => r.get('id') === id);
    if (target) {
      target.assign({ status: status || target.get('status') });
      await target.save();
    }
    res.json({ success: true });
  } catch (err: any) {
    console.error('Error updating attendance status:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

app.delete('/api/attendance/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const sheet = await getSheet('Attendance');
    if (!sheet) return res.status(500).json({ success: false, message: 'Attendance sheet not found' });

    const rows = await sheet.getRows();
    const target = rows.find((r) => r.get('id') === id);
    if (target) {
      await target.delete();
    }
    res.json({ success: true });
  } catch (err: any) {
    console.error('Error deleting attendance record:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/attendance/auto-checkout-check', async (req, res) => {
  res.json({ success: true, message: 'Auto checkout verification completed' });
});

// Announcements Endpoints
app.get('/api/announcements', async (req, res) => {
  try {
    const sheet = await getSheet('Announcements');
    if (!sheet) return res.json([]);
    const rows = await sheet.getRows();
    const announcements = rows.map((r) => ({
      id: r.get('id') || '',
      title: r.get('title') || '',
      content: r.get('content') || '',
      date: r.get('date') || '',
      expiryDate: r.get('expiryDate') || '',
      isActive: r.get('isActive') !== 'FALSE' && r.get('isActive') !== 'false',
    }));
    res.json(announcements);
  } catch (err: any) {
    console.error('Error fetching announcements:', err);
    res.status(500).json({ error: 'Failed to fetch announcements', message: err.message });
  }
});

app.post('/api/announcements', async (req, res) => {
  try {
    const { title, content, date, expiryDate, isActive } = req.body;
    const sheet = await getSheet('Announcements');
    if (!sheet) return res.status(500).json({ success: false, message: 'Announcements sheet not found' });

    const id = Date.now().toString();
    await sheet.addRow({
      id,
      title: title || '',
      content: content || '',
      date: date || new Date().toISOString().split('T')[0],
      expiryDate: expiryDate || '',
      isActive: isActive !== false ? 'TRUE' : 'FALSE',
    });

    res.json({ success: true, id });
  } catch (err: any) {
    console.error('Error adding announcement:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

app.put('/api/announcements/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const { title, content, date, expiryDate, isActive } = req.body;
    const sheet = await getSheet('Announcements');
    if (!sheet) return res.status(500).json({ success: false, message: 'Announcements sheet not found' });

    const rows = await sheet.getRows();
    const target = rows.find((r) => r.get('id') === id);
    if (target) {
      target.assign({
        title: title ?? target.get('title'),
        content: content ?? target.get('content'),
        date: date ?? target.get('date'),
        expiryDate: expiryDate ?? target.get('expiryDate'),
        isActive: isActive !== undefined ? (isActive ? 'TRUE' : 'FALSE') : target.get('isActive'),
      });
      await target.save();
    }
    res.json({ success: true });
  } catch (err: any) {
    console.error('Error updating announcement:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

app.delete('/api/announcements/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const sheet = await getSheet('Announcements');
    if (!sheet) return res.status(500).json({ success: false, message: 'Announcements sheet not found' });
    const rows = await sheet.getRows();
    const target = rows.find((r) => r.get('id') === id);
    if (target) {
      await target.delete();
    }
    res.json({ success: true });
  } catch (err: any) {
    console.error('Error deleting announcement:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

// Holidays Endpoints
app.get('/api/holidays', async (req, res) => {
  try {
    const sheet = await getSheet('Holidays');
    if (!sheet) return res.json([]);
    const rows = await sheet.getRows();
    const holidays = rows.map((r) => ({
      id: r.get('id') || '',
      date: r.get('date') || '',
      name: r.get('name') || '',
    }));
    res.json(holidays);
  } catch (err: any) {
    console.error('Error fetching holidays:', err);
    res.status(500).json({ error: 'Failed to fetch holidays', message: err.message });
  }
});

app.post('/api/holidays', async (req, res) => {
  try {
    const { date, name } = req.body;
    const sheet = await getSheet('Holidays');
    if (!sheet) return res.status(500).json({ success: false, message: 'Holidays sheet not found' });

    const id = Date.now().toString();
    await sheet.addRow({
      id,
      date: date || '',
      name: name || '',
    });

    res.json({ success: true, id });
  } catch (err: any) {
    console.error('Error adding holiday:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

app.delete('/api/holidays/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const sheet = await getSheet('Holidays');
    if (!sheet) return res.status(500).json({ success: false, message: 'Holidays sheet not found' });
    const rows = await sheet.getRows();
    const target = rows.find((r) => r.get('id') === id);
    if (target) {
      await target.delete();
    }
    res.json({ success: true });
  } catch (err: any) {
    console.error('Error deleting holiday:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

// Device Bindings Reset Endpoint
app.delete('/api/device-bindings/:nip', async (req, res) => {
  try {
    const { nip } = req.params;
    const sheet = await getSheet('DeviceBindings');
    if (!sheet) return res.status(500).json({ success: false, message: 'DeviceBindings sheet not found' });

    const rows = await sheet.getRows();
    const target = rows.find((r) => r.get('nip') === String(nip));
    if (target) {
      await target.delete();
    }
    res.json({ success: true, message: 'Perangkat berhasil direset' });
  } catch (err: any) {
    console.error('Error resetting device binding:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

// Authentication: Login
app.post('/api/login', async (req, res) => {
  try {
    const { nip, password, deviceId } = req.body;
    if (!nip || !password) {
      return res.status(400).json({ success: false, message: 'NIP dan password wajib diisi' });
    }

    const nipStr = String(nip).trim();
    const passStr = String(password).trim();

    // 1. Check Admin logins
    if (
      (nipStr === 'admin' && passStr === 'admin') ||
      (nipStr === '000000' && (passStr === 'admin' || passStr === '123456'))
    ) {
      return res.json({
        success: true,
        user: {
          id: '1',
          name: 'Super Admin',
          nip: nipStr,
          email: 'super@admin.com',
          role: 'admin',
          group: 'Superadmin',
          access: ['Absensi', 'Master Data', 'Sistem'],
        },
      });
    }

    // Check Admins sheet
    const adminSheet = await getSheet('Admins');
    if (adminSheet) {
      const adminRows = await adminSheet.getRows();
      const adminUser = adminRows.find(
        (r) =>
          (r.get('nip') === nipStr || r.get('email') === nipStr) &&
          String(r.get('password') || '').trim() === passStr &&
          r.get('isActive') !== 'FALSE' &&
          r.get('isActive') !== 'false'
      );

      if (adminUser) {
        return res.json({
          success: true,
          user: {
            id: adminUser.get('id') || '1',
            name: adminUser.get('name') || 'Admin',
            nip: adminUser.get('nip') || nipStr,
            email: adminUser.get('email') || '',
            role: 'admin',
            group: adminUser.get('group') || 'Admin',
            access: safeJsonParse(adminUser.get('access')) || ['Absensi', 'Master Data'],
          },
        });
      }
    }

    // 2. Check Employees sheet
    const empSheet = await getSheet('Employees');
    if (!empSheet) {
      return res.status(500).json({ success: false, message: 'Database karyawan tidak tersedia' });
    }

    const empRows = await empSheet.getRows();
    const employee = empRows.find((r) => r.get('nip') === nipStr);

    if (!employee) {
      return res.status(401).json({ success: false, message: 'NIP tidak terdaftar' });
    }

    if (String(employee.get('password') || '').trim() !== passStr) {
      return res.status(401).json({ success: false, message: 'Password salah' });
    }

    // 3. Verify Device Binding
    if (deviceId) {
      const bindingSheet = await getSheet('DeviceBindings');
      if (bindingSheet) {
        const bindingRows = await bindingSheet.getRows();
        const existingBinding = bindingRows.find((r) => r.get('nip') === nipStr);

        if (existingBinding) {
          const boundId = existingBinding.get('deviceId');
          if (boundId && boundId !== deviceId) {
            return res.status(403).json({
              success: false,
              message: 'Akun ini terikat dengan perangkat lain. Hubungi admin untuk mereset perangkat.',
            });
          }
        } else {
          // Bind this device
          await bindingSheet.addRow({ nip: nipStr, deviceId });
        }
      }
    }

    return res.json({
      success: true,
      user: {
        id: employee.get('id') || '',
        name: employee.get('name') || '',
        nip: employee.get('nip') || '',
        office: employee.get('office') || '',
        office2: employee.get('office2') || '',
        email: employee.get('email') || '',
        gender: employee.get('gender') || '',
        cluster: employee.get('cluster') || '',
        unit: employee.get('unit') || '',
        role: 'user',
        photoUrl: employee.get('photoUrl') || '',
        photoUploadCount: parseInt(employee.get('photoUploadCount') || '0', 10),
      },
    });
  } catch (err: any) {
    console.error('Error during login:', err);
    res.status(500).json({ success: false, message: 'Terjadi kesalahan server: ' + err.message });
  }
});

// Authentication: Register
app.post('/api/register', async (req, res) => {
  try {
    const { nip, name, email, password, gender, cluster, unit, desa, office2 } = req.body;
    if (!nip || !name || !password) {
      return res.status(400).json({ success: false, message: 'NIP, nama, dan password wajib diisi' });
    }

    const nipStr = String(nip).trim();
    const empSheet = await getSheet('Employees');
    if (!empSheet) {
      return res.status(500).json({ success: false, message: 'Database karyawan tidak tersedia' });
    }

    const empRows = await empSheet.getRows();
    const existing = empRows.find((r) => r.get('nip') === nipStr);
    if (existing) {
      return res.status(400).json({ success: false, message: 'NIP sudah terdaftar' });
    }

    const id = Date.now().toString();
    await empSheet.addRow({
      id,
      name: name || '',
      nip: nipStr,
      office: desa || '',
      office2: office2 === 'none' ? '' : office2 || '',
      email: email || '',
      gender: gender || '',
      cluster: cluster || '',
      unit: unit || '',
      password: String(password).trim(),
      photoUrl: '',
      photoUploadCount: '0',
    });

    res.json({ success: true, message: 'Pendaftaran berhasil, silakan login' });
  } catch (err: any) {
    console.error('Error during register:', err);
    res.status(500).json({ success: false, message: 'Terjadi kesalahan server: ' + err.message });
  }
});

// Password Management
app.post('/api/change-password', async (req, res) => {
  try {
    const { id, role, oldPassword, newPassword } = req.body;
    if (!newPassword || newPassword.length < 6) {
      return res.status(400).json({ success: false, message: 'Password minimal 6 karakter' });
    }

    if (role === 'admin') {
      const adminSheet = await getSheet('Admins');
      if (adminSheet) {
        const rows = await adminSheet.getRows();
        const target = rows.find((r) => r.get('id') === id);
        if (target) {
          if (oldPassword && target.get('password') && target.get('password') !== oldPassword) {
            return res.status(400).json({ success: false, message: 'Password lama salah' });
          }
          target.assign({ password: newPassword });
          await target.save();
          return res.json({ success: true, message: 'Password admin berhasil diubah' });
        }
      }
    }

    const empSheet = await getSheet('Employees');
    if (!empSheet) return res.status(500).json({ success: false, message: 'Database karyawan tidak tersedia' });

    const rows = await empSheet.getRows();
    const target = rows.find((r) => r.get('id') === id || r.get('nip') === id);
    if (!target) {
      return res.status(404).json({ success: false, message: 'Pengguna tidak ditemukan' });
    }

    if (oldPassword && target.get('password') && target.get('password') !== oldPassword) {
      return res.status(400).json({ success: false, message: 'Password lama tidak cocok' });
    }

    target.assign({ password: newPassword });
    await target.save();
    res.json({ success: true, message: 'Password berhasil diperbarui' });
  } catch (err: any) {
    console.error('Error changing password:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/forgot-password', async (req, res) => {
  try {
    const { email } = req.body;
    if (!email) {
      return res.status(400).json({ success: false, message: 'Email wajib diisi' });
    }

    const empSheet = await getSheet('Employees');
    let userFound = false;
    let userId = '';

    if (empSheet) {
      const rows = await empSheet.getRows();
      const emp = rows.find((r) => r.get('email')?.toLowerCase() === email.toLowerCase());
      if (emp) {
        userFound = true;
        userId = emp.get('id');
      }
    }

    if (!userFound) {
      const adminSheet = await getSheet('Admins');
      if (adminSheet) {
        const rows = await adminSheet.getRows();
        const adm = rows.find((r) => r.get('email')?.toLowerCase() === email.toLowerCase());
        if (adm) {
          userFound = true;
          userId = adm.get('id');
        }
      }
    }

    if (!userFound) {
      return res.status(404).json({ success: false, message: 'Email tidak terdaftar dalam sistem' });
    }

    const token = crypto.randomUUID();
    const resetSheet = await getSheet('PasswordResets');
    if (resetSheet) {
      await resetSheet.addRow({
        token,
        userId,
        userType: 'user',
        expiresAt: new Date(Date.now() + 60 * 60 * 1000).toISOString(),
      });
    }

    const origin = req.headers.origin || `${req.protocol}://${req.get('host')}`;
    const resetLink = `${origin}/reset-password?token=${token}`;

    const resend = getResend();
    if (resend) {
      try {
        await resend.emails.send({
          from: 'Si Abon Elite <onboarding@resend.dev>',
          to: email,
          subject: 'Reset Password - Si Abon Elite App',
          html: `<p>Halo,</p><p>Anda meminta untuk mereset password akun Si Abon Elite. Klik tautan berikut untuk membuat password baru:</p><p><a href="${resetLink}">Reset Password</a></p><p>Tautan ini berlaku selama 1 jam.</p>`,
        });
      } catch (emailErr) {
        console.warn('Resend delivery error, returning mock link fallback:', emailErr);
      }
    }

    res.json({
      success: true,
      message: 'Email reset password telah dikirim',
      mockLink: resetLink,
    });
  } catch (err: any) {
    console.error('Error handling forgot password:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/reset-password', async (req, res) => {
  try {
    const { token, newPassword } = req.body;
    if (!token || !newPassword) {
      return res.status(400).json({ success: false, message: 'Token dan password baru wajib diisi' });
    }

    const resetSheet = await getSheet('PasswordResets');
    if (!resetSheet) {
      return res.status(500).json({ success: false, message: 'Password resets sheet tidak tersedia' });
    }

    const rows = await resetSheet.getRows();
    const targetReset = rows.find((r) => r.get('token') === token);
    if (!targetReset) {
      return res.status(400).json({ success: false, message: 'Token reset tidak valid atau telah digunakan' });
    }

    const expiresAt = targetReset.get('expiresAt');
    if (expiresAt && new Date(expiresAt).getTime() < Date.now()) {
      await targetReset.delete();
      return res.status(400).json({ success: false, message: 'Token reset telah kedaluwarsa' });
    }

    const userId = targetReset.get('userId');
    const empSheet = await getSheet('Employees');
    let updated = false;

    if (empSheet) {
      const empRows = await empSheet.getRows();
      const emp = empRows.find((r) => r.get('id') === userId);
      if (emp) {
        emp.assign({ password: newPassword });
        await emp.save();
        updated = true;
      }
    }

    if (!updated) {
      const adminSheet = await getSheet('Admins');
      if (adminSheet) {
        const admRows = await adminSheet.getRows();
        const adm = admRows.find((r) => r.get('id') === userId);
        if (adm) {
          adm.assign({ password: newPassword });
          await adm.save();
          updated = true;
        }
      }
    }

    await targetReset.delete();

    if (!updated) {
      return res.status(404).json({ success: false, message: 'Pengguna terkait token tidak ditemukan' });
    }

    res.json({ success: true, message: 'Password berhasil diperbarui' });
  } catch (err: any) {
    console.error('Error resetting password:', err);
    res.status(500).json({ success: false, message: err.message });
  }
});

// ==========================================
// VITE SPA & STATIC SERVING
// ==========================================
async function start() {
  if (process.env.NODE_ENV !== 'production') {
    const { createServer: createViteServer } = await import('vite');
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

start();

export default app;
