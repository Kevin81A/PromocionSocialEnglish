// ========================================================
// VerbFlow - Supabase Client Configuration
// Manejo centralizado de credenciales y cliente Supabase
// I.E. Promoción Social • Palermo, Huila
// ========================================================

const SupabaseConfig = (() => {
  const config = {
    url: (typeof window !== "undefined" && window.__ENV__ && window.__ENV__.SUPABASE_URL) 
      || (typeof localStorage !== "undefined" && localStorage.getItem("vf_supabase_url"))
      || "https://tu-proyecto.supabase.co",
    anonKey: (typeof window !== "undefined" && window.__ENV__ && window.__ENV__.SUPABASE_ANON_KEY)
      || (typeof localStorage !== "undefined" && localStorage.getItem("vf_supabase_key"))
      || "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.dummy_anon_key"
  };

  function isConfigured() {
    return Boolean(
      config.url && 
      config.anonKey && 
      !config.url.includes("tu-proyecto.supabase.co") && 
      !config.anonKey.includes("dummy_anon_key")
    );
  }

  function getClient() {
    if (typeof window !== "undefined" && window.supabase && isConfigured()) {
      if (!window.__supabaseInstance) {
        window.__supabaseInstance = window.supabase.createClient(config.url, config.anonKey, {
          auth: {
            persistSession: true,
            autoRefreshToken: true,
            detectSessionInUrl: true
          }
        });
      }
      return window.__supabaseInstance;
    }
    return null;
  }

  return {
    config,
    isConfigured,
    getClient
  };
})();

if (typeof window !== "undefined") {
  window.SupabaseConfig = SupabaseConfig;
}
