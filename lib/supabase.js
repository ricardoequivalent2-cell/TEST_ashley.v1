"use client";
import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = "https://efqbcqlrxwgygobvlvcs.supabase.co";
const SUPABASE_ANON = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImVmcWJjcWxyeHdneWdvYnZsdmNzIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODM3NjU2NjksImV4cCI6MjA5OTM0MTY2OX0.nLp06dHBZBUzbze5_cFT3WQjeSq-f7OeGrUhJEvXVu8";

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON);
