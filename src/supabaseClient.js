import { createClient } from "@supabase/supabase-js";

// These come from Vite environment variables.
// Locally: put them in a file named .env.local (see .env.example)
// On GitHub: set them as repo Secrets, injected during the Actions build (see workflow file)
const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseAnonKey = import.meta.env.VITE_SUPABASE_ANON_KEY;

if (!supabaseUrl || !supabaseAnonKey) {
  console.warn(
    "Supabase URL/Key are missing. Set VITE_SUPABASE_URL and VITE_SUPABASE_ANON_KEY."
  );
}

export const supabase = createClient(supabaseUrl, supabaseAnonKey);

export const BOOKS_BUCKET = "books";
export const ASSETS_BUCKET = "assets";
