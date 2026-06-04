/* ============================================================
   VCSAR Team 1 — Supabase configuration (shared)
   Used by team.html (read roster) and admin.html (manage roster).

   Paste the two values from your Supabase project below
   (Project Settings ▸ API). See form-backend/TEAM-ADMIN-SETUP.txt.
   Leave them blank to run in DEMO mode (no backend).
   These keys are safe to expose publicly: the anon key only allows
   what your Row Level Security policies permit (public read; edits
   require login).
   ============================================================ */
window.SUPABASE_URL  = "https://rhnxpcjlmdzialxxsrpa.supabase.co";   // e.g. "https://abcdefgh.supabase.co"
window.SUPABASE_ANON = "sb_publishable_s6tqiGb34P9pNFkinVbnrA_ex-WV_NG";   // e.g. "eyJhbGciOiJIUzI1NiIsInR5cCI6..."

// Storage bucket that holds member photos (created during setup).
window.SUPABASE_BUCKET = "member-photos";
