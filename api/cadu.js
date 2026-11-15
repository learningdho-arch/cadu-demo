module.exports = async function handler(req, res) {
 // Hanya menerima POST
 if (req.method !== 'POST') {
 return res.status(405).json({
 success: false,
 error: 'Method not allowed'
 });
 }

 // Ambil konfigurasi rahasia dari Vercel
 const scriptUrl = process.env.CADU_SCRIPT_URL;
 const secret = process.env.CADU_API_SECRET;

 if (!scriptUrl || !secret) {
 return res.status(500).json({
 success: false,
 error: 'Konfigurasi server belum lengkap'
 });
 }

 try {
 // Baca request dari frontend
 const body = typeof req.body === 'string'
 ? JSON.parse(req.body)
 : (req.body || {});

 if (!body.action || typeof body.action !== 'string') {
 return res.status(400).json({
 success: false,
 error: 'Action tidak valid'
 });
 }

 // Kirim request ke Google Apps Script
 const response = await fetch(scriptUrl, {
 method: 'POST',
 headers: {
 'Content-Type': 'application/json'
 },
 body: JSON.stringify({
 action: body.action,
 args: Array.isArray(body.args) ? body.args : [],
 secret: secret
 }),
 redirect: 'follow'
 });

 const responseText = await response.text();

 let result;

 try {
 result = JSON.parse(responseText);
 } catch {
 return res.status(502).json({
 success: false,
 error: 'Respons Apps Script bukan JSON yang valid'
 });
 }

 return res.status(200).json(result);

 } catch (error) {
 return res.status(502).json({
 success: false,
 error: 'Gagal menghubungkan ke Apps Script'
 });
 }
};
