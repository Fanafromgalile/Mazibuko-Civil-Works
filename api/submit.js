const MAX_FIELD_LENGTH = 5000;
const FIELDS = ['type', 'fullName', 'email', 'contactInfo', 'location', 'service', 'details', 'question', 'topic'];
const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

module.exports = async function handler(req, res) {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(204).end();
  }

  if (req.method !== 'POST') {
    res.setHeader('Allow', 'POST, OPTIONS');
    return res.status(405).json({ status: 'error', message: 'Method not allowed.' });
  }

  const scriptUrl = process.env.next_public_apps_script_url || process.env.APPS_SCRIPT_URL;
  if (!scriptUrl) {
    return res.status(500).json({ status: 'error', message: 'Submission service is not configured.' });
  }

  let body = req.body;
  if (typeof body === 'string') {
    try {
      body = JSON.parse(body);
    } catch {
      return res.status(400).json({ status: 'error', message: 'Invalid JSON payload.' });
    }
  }
  body = body || {};

  const payload = {};
  for (const field of FIELDS) {
    payload[field] = String(body[field] ?? '').trim().slice(0, MAX_FIELD_LENGTH);
  }

  if (payload.type !== 'quote' && payload.type !== 'faq') {
    return res.status(400).json({ status: 'error', message: 'Invalid submission type.' });
  }
  if (!payload.fullName) {
    return res.status(400).json({ status: 'error', message: 'Please provide your full name.' });
  }
  if (!EMAIL_PATTERN.test(payload.email)) {
    return res.status(400).json({ status: 'error', message: 'Please provide a valid email address.' });
  }

  try {
    const upstream = await fetch(scriptUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload),
      redirect: 'follow'
    });
    const text = await upstream.text();
    let result;
    try {
      result = JSON.parse(text);
    } catch {
      console.error('Apps Script returned a non-JSON response', upstream.status);
      return res.status(502).json({ status: 'error', message: 'Submission service returned an unexpected response.' });
    }
    return res.status(result.status === 'success' ? 200 : 400).json(result);
  } catch (error) {
    console.error('Failed to reach Apps Script', error);
    return res.status(502).json({ status: 'error', message: 'Could not reach the submission service.' });
  }
};
