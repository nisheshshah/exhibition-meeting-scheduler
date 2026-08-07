-- ====================================================================
-- UNIVERSAL OLEORESINS - EXHIBITION MEETING SCHEDULER (SUPABASE SCHEMA)
-- Fully Idempotent & Production-Ready Script
-- ====================================================================

-- 1. Create Salesmen Table
CREATE TABLE IF NOT EXISTS public.salesmen (
    id VARCHAR(20) PRIMARY KEY,
    name VARCHAR(100) NOT NULL,
    email VARCHAR(150) NOT NULL,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL
);

-- Insert Default Sales Team
INSERT INTO public.salesmen (id, name, email) VALUES
('S002', 'Jai Shah', 'jaishah@universaloleoresins.com'),
('S003', 'Nishesh Shah', 'nisheshshah@universaloleoresins.com'),
('S005', 'Paul Thampy', 'intsales@universaloleoresins.com'),
('S006', 'Bikash Kar', 'domsaleseast@universaloleoresins.com'),
('S007', 'Harshita Shah', 'harshitashah@universaloleoresins.com'),
('S009', 'Payal', 'domsales@universaloleoresins.com'),
('S010', 'Machindranath', 'domsalesmh@universaloleoresins.com'),
('S011', 'Shishir Shah', 'shishirshah@xtractiva.com'),
('S013', 'Saurabh', 'domsalesgj@universaloleoresins.com'),
('S014', 'Kiruthi Kumar', 'domsalessouth@universaloleoresins.com'),
('S015', 'Shubham', 'horecawest@universaloleoresins.com'),
('S020', 'Jigesh Shah', 'jigeshshah@universaloleoresins.com')
ON CONFLICT (id) DO UPDATE SET name = EXCLUDED.name, email = EXCLUDED.email;

-- 2. Create Exhibitions Table FIRST before foreign keys
CREATE TABLE IF NOT EXISTS public.exhibitions (
    id VARCHAR(50) PRIMARY KEY,
    title VARCHAR(150) NOT NULL,
    venue VARCHAR(200) NOT NULL,
    location VARCHAR(100) NOT NULL,
    start_date DATE NOT NULL,
    end_date DATE NOT NULL,
    slot_length_minutes INT DEFAULT 15,
    timezone VARCHAR(50) DEFAULT 'Asia/Kolkata',
    is_active BOOLEAN DEFAULT TRUE,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL
);

-- Idempotent Column Additions for Upgrade Schema
ALTER TABLE public.exhibitions ADD COLUMN IF NOT EXISTS timezone VARCHAR(50) DEFAULT 'Asia/Kolkata';

-- Insert Default Exhibition & Dev Staging Exhibition
INSERT INTO public.exhibitions (id, title, venue, location, start_date, end_date, slot_length_minutes, timezone, is_active) VALUES
('fi-india-2026', 'Fi India 2026', '(BEC), Goregaon, Mumbai', 'Stall 3D38, Hall 3', '2026-08-26', '2026-08-28', 15, 'Asia/Kolkata', TRUE),
('fi-india-2026-dev', 'Fi India 2026 (Staging / Dev Test)', '(BEC), Goregaon, Mumbai', 'Stall 3D38 (Staging Test)', '2026-08-26', '2026-08-28', 15, 'Asia/Kolkata', TRUE)
ON CONFLICT (id) DO UPDATE SET title = EXCLUDED.title, venue = EXCLUDED.venue, location = EXCLUDED.location, timezone = EXCLUDED.timezone;

-- 3. Create Exhibition Team Table (Foreign Keys to exhibitions & salesmen)
CREATE TABLE IF NOT EXISTS public.exhibition_team (
    exhibition_id VARCHAR(50) NOT NULL,
    salesman_id VARCHAR(20) NOT NULL,
    CONSTRAINT pk_exhibition_team PRIMARY KEY (exhibition_id, salesman_id),
    CONSTRAINT fk_exhibition FOREIGN KEY (exhibition_id) REFERENCES public.exhibitions(id) ON DELETE CASCADE,
    CONSTRAINT fk_salesman FOREIGN KEY (salesman_id) REFERENCES public.salesmen(id) ON DELETE CASCADE
);

-- Assign all salesmen to Fi India 2026 & Staging
INSERT INTO public.exhibition_team (exhibition_id, salesman_id)
SELECT 'fi-india-2026', id FROM public.salesmen
ON CONFLICT DO NOTHING;

INSERT INTO public.exhibition_team (exhibition_id, salesman_id)
SELECT 'fi-india-2026-dev', id FROM public.salesmen
ON CONFLICT DO NOTHING;

-- 4. Create Bookings Table
CREATE TABLE IF NOT EXISTS public.bookings (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    ref VARCHAR(20) UNIQUE NOT NULL,
    exhibition_id VARCHAR(50) NOT NULL,
    salesman_id VARCHAR(20) NOT NULL,
    date DATE NOT NULL,
    time VARCHAR(10) NOT NULL,
    first_name VARCHAR(100) NOT NULL,
    last_name VARCHAR(100) NOT NULL,
    company VARCHAR(150) NOT NULL,
    email VARCHAR(150) NOT NULL,
    phone VARCHAR(50) NOT NULL,
    status VARCHAR(20) DEFAULT 'confirmed',
    lead_tier VARCHAR(10) DEFAULT 'UNASSIGNED',
    product_interests TEXT DEFAULT '',
    rep_notes TEXT DEFAULT '',
    checked_in BOOLEAN DEFAULT FALSE,
    checked_in_at TIMESTAMP WITH TIME ZONE,
    created_at TIMESTAMP WITH TIME ZONE DEFAULT timezone('utc'::text, now()) NOT NULL,
    
    CONSTRAINT fk_booking_exhibition FOREIGN KEY (exhibition_id) REFERENCES public.exhibitions(id) ON DELETE CASCADE,
    CONSTRAINT fk_booking_salesman FOREIGN KEY (salesman_id) REFERENCES public.salesmen(id) ON DELETE CASCADE,
    CONSTRAINT unique_salesman_slot UNIQUE (exhibition_id, salesman_id, date, time)
);

-- Idempotent column additions for existing bookings tables
ALTER TABLE public.bookings ADD COLUMN IF NOT EXISTS lead_tier VARCHAR(10) DEFAULT 'UNASSIGNED';
ALTER TABLE public.bookings ADD COLUMN IF NOT EXISTS product_interests TEXT DEFAULT '';
ALTER TABLE public.bookings ADD COLUMN IF NOT EXISTS rep_notes TEXT DEFAULT '';
ALTER TABLE public.bookings ADD COLUMN IF NOT EXISTS checked_in BOOLEAN DEFAULT FALSE;
ALTER TABLE public.bookings ADD COLUMN IF NOT EXISTS checked_in_at TIMESTAMP WITH TIME ZONE;

-- 5. Safely Enable Supabase Realtime
DO $$
BEGIN
  IF NOT EXISTS (
    SELECT 1 FROM pg_publication_tables 
    WHERE pubname = 'supabase_realtime' AND schemaname = 'public' AND tablename = 'bookings'
  ) THEN
    ALTER PUBLICATION supabase_realtime ADD TABLE public.bookings;
  END IF;
END $$;

-- 6. Configure Row Level Security (RLS) & Policies cleanly
ALTER TABLE public.salesmen ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.exhibitions ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.exhibition_team ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.bookings ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "Allow public read access on salesmen" ON public.salesmen;
DROP POLICY IF EXISTS "Allow public read access on exhibitions" ON public.exhibitions;
DROP POLICY IF EXISTS "Allow public read access on exhibition_team" ON public.exhibition_team;
DROP POLICY IF EXISTS "Allow public read access on bookings" ON public.bookings;
DROP POLICY IF EXISTS "Allow public insert on bookings" ON public.bookings;
DROP POLICY IF EXISTS "Allow public update on bookings" ON public.bookings;

CREATE POLICY "Allow public read access on salesmen" ON public.salesmen FOR SELECT USING (true);
CREATE POLICY "Allow public read access on exhibitions" ON public.exhibitions FOR SELECT USING (true);
CREATE POLICY "Allow public read access on exhibition_team" ON public.exhibition_team FOR SELECT USING (true);
CREATE POLICY "Allow public read access on bookings" ON public.bookings FOR SELECT USING (true);
CREATE POLICY "Allow public insert on bookings" ON public.bookings FOR INSERT WITH CHECK (true);
CREATE POLICY "Allow public update on bookings" ON public.bookings FOR UPDATE USING (true);
