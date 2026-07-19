import { supabase } from "./supabaseClient";

// Mesma "forma" da API usada antes (window.storage do Claude), mas
// agora gravando de verdade na tabela `kv_store` do Supabase.
// get(key)        -> { key, value } | null
// set(key, value) -> { key, value } | null
export const storage = {
  async get(key) {
    const { data, error } = await supabase
      .from("kv_store")
      .select("value")
      .eq("key", key)
      .maybeSingle();

    if (error) {
      console.error("storage.get error:", error);
      throw error;
    }
    if (!data) return null;
    return { key, value: data.value };
  },

  async set(key, value) {
    const { error } = await supabase
      .from("kv_store")
      .upsert({ key, value, updated_at: new Date().toISOString() }, { onConflict: "key" });

    if (error) {
      console.error("storage.set error:", error);
      return null;
    }
    return { key, value };
  },
};
