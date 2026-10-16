# DIIT 21st CSE Closing Ceremony

Convocation-style registration website for DIIT's 21st CSE batch. Private organizer/admin dashboard, manual payment approval, automatic fee calculation, email confirmations and QR ticket verification.

## Setup

1. Install Node.js 22 or newer.
2. Run `npm install`.
3. Copy `.env.example` to `.env` and supply a strong `ADMIN_PASSWORD`, a different random `SESSION_SECRET`, and SMTP credentials (`SMTP_HOST`, `SMTP_USER`, `SMTP_PASS`, `MAIL_FROM`).
4. Run `npm start` and visit `http://localhost:3000`. The admin page is `/admin`.

## Event

- Deadline: 27 October 2026 (Bangladesh time)
- Ceremony: after final defence, exact date to be announced
- Institute: https://diit.edu.bd/
- Venue: https://greengarden.com.bd/
- Student fee: ৳1,000; guest fee: ৳500 each
- Cash or manually verified bKash: 01940372926
- Account: 1073055620001
- Contact: Al-Rafi Chowdhury, 01940372926

## Production

Deploy on a Node.js hosting service **with persistent SQLite storage**. GitHub Pages alone will not run the server. Configure `NODE_ENV=production`, `PUBLIC_BASE_URL=https://your-domain`, secure admin credentials, SMTP service, and `TRUST_PROXY=1` behind a trusted reverse proxy. Test the complete registration, payment verification and email workflow before accepting real attendees. Keep `.env` and database files private.

Admin payment confirmation is **manual**: bKash integration does not automatically check transactions. Email and QR tickets are sent only after the organizer verifies the payment. QR codes link to a server-hosted verification page; they do not directly embed contact information.

This is a starter implementation and has not yet been production-tested or deployed.
