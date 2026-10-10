-- =====================================================================
-- منظومة التخطيط التشغيلي والموازنة — مخطط قاعدة البيانات (PostgreSQL 16)
-- مشتق من ملف Excel: DATA / كشاف الأهداف والمؤشرات / كشاف المشاريع / مشروع
-- =====================================================================
CREATE EXTENSION IF NOT EXISTS pgcrypto;
CREATE EXTENSION IF NOT EXISTS citext;

-- ---------- البيانات المرجعية ----------
CREATE TABLE perspectives (            -- بادئة رمز الهدف: س ص م ل ت
  id          char(1) PRIMARY KEY,
  name        text NOT NULL
);

CREATE TABLE objectives (              -- كشاف الأهداف والمؤشرات D:E
  id          text PRIMARY KEY,        -- 'س1'
  perspective char(1) NOT NULL REFERENCES perspectives(id),
  name        text NOT NULL UNIQUE,
  is_active   boolean NOT NULL DEFAULT true
);

CREATE TYPE kpi_unit AS ENUM ('عدد','نسبة مئوية','ريال','مستوى','أخرى');
CREATE TABLE kpis (                    -- كشاف الأهداف والمؤشرات F:G
  id           text PRIMARY KEY,       -- 'س1-1'
  objective_id text NOT NULL REFERENCES objectives(id),
  name         text NOT NULL,
  unit         kpi_unit NOT NULL,
  is_active    boolean NOT NULL DEFAULT true,
  UNIQUE (objective_id, name)
);

CREATE TABLE departments (             -- DATA!H6:H17
  id        text PRIMARY KEY,          -- 'D01'
  name      text NOT NULL UNIQUE,
  is_active boolean NOT NULL DEFAULT true
);

CREATE TABLE sections (                -- DATA!J:K (كانت قيماً افتراضية — تُعبأ لاحقاً)
  id            text PRIMARY KEY,
  department_id text NOT NULL REFERENCES departments(id),
  name          text NOT NULL,
  UNIQUE (department_id, name)
);

CREATE TABLE plan_years (
  year         smallint PRIMARY KEY,   -- 2026
  week1_start  date NOT NULL,          -- أول سبت في السنة (بديل تقويم 2024 الخاطئ)
  is_locked    boolean NOT NULL DEFAULT false
);

-- ---------- المستخدمون والصلاحيات ----------
CREATE TYPE app_role AS ENUM ('admin','planner','dept_editor','viewer');
CREATE TABLE users (
  id            uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  sso_subject   text UNIQUE NOT NULL,  -- معرف Entra ID / Google
  display_name  text NOT NULL,
  email         citext UNIQUE,
  role          app_role NOT NULL DEFAULT 'viewer',
  department_id text REFERENCES departments(id),  -- نطاق dept_editor
  created_at    timestamptz NOT NULL DEFAULT now()
);

-- ---------- المشاريع (كشاف المشاريع + رأس ورقة مشروع) ----------
CREATE TABLE projects (
  id            uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  code          text UNIQUE NOT NULL,                 -- P-2026-001
  plan_year     smallint NOT NULL REFERENCES plan_years(year),
  name          text NOT NULL CHECK (length(name) >= 3),
  description   text,
  objective_id  text NOT NULL REFERENCES objectives(id),
  kpi_id        text NOT NULL REFERENCES kpis(id),
  department_id text NOT NULL REFERENCES departments(id),
  section       text,
  branch        text,
  priority      smallint NOT NULL CHECK (priority BETWEEN 1 AND 3),
  manager       text,
  budget        numeric(14,2) NOT NULL DEFAULT 0 CHECK (budget >= 0),
  manual_progress numeric(5,2) CHECK (manual_progress BETWEEN 0 AND 100),
  version       integer NOT NULL DEFAULT 1,           -- قفل تفاؤلي
  created_by    uuid REFERENCES users(id),
  created_at    timestamptz NOT NULL DEFAULT now(),
  updated_by    uuid REFERENCES users(id),
  updated_at    timestamptz NOT NULL DEFAULT now(),
  deleted_at    timestamptz                           -- حذف منطقي
);
-- المؤشر يجب أن يتبع الهدف المختار (يحل محل قائمة OFFSET المتتالية)
CREATE UNIQUE INDEX kpis_obj_pair ON kpis(id, objective_id);
ALTER TABLE projects ADD CONSTRAINT projects_kpi_matches_objective
  FOREIGN KEY (kpi_id, objective_id) REFERENCES kpis(id, objective_id);

CREATE TABLE project_monthly_budget (  -- كشاف المشاريع H:S
  project_id uuid REFERENCES projects(id) ON DELETE CASCADE,
  month      smallint CHECK (month BETWEEN 1 AND 12),
  amount     numeric(14,2) NOT NULL DEFAULT 0 CHECK (amount >= 0),
  PRIMARY KEY (project_id, month)
);

CREATE TABLE project_steps (           -- ورقة مشروع A12:J30
  id            uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  project_id    uuid NOT NULL REFERENCES projects(id) ON DELETE CASCADE,
  seq           smallint NOT NULL,
  name          text NOT NULL,
  needs         text,                  -- بشرية | مرافق | أجهزة | مطبوعات
  owner         text,
  cost          numeric(14,2) NOT NULL DEFAULT 0 CHECK (cost >= 0),
  plan_start    smallint NOT NULL CHECK (plan_start BETWEEN 1 AND 52),
  plan_weeks    smallint NOT NULL CHECK (plan_weeks >= 1),
  actual_start  smallint CHECK (actual_start BETWEEN 1 AND 52),
  actual_weeks  smallint CHECK (actual_weeks >= 0),
  pct           numeric(5,2) NOT NULL DEFAULT 0 CHECK (pct BETWEEN 0 AND 100),
  CHECK (plan_start + plan_weeks - 1 <= 52),
  CHECK (actual_start IS NULL OR actual_start + coalesce(actual_weeks,0) - 1 <= 52),
  UNIQUE (project_id, seq)
);

-- ---------- سجل التدقيق (إلحاقي فقط) ----------
CREATE TABLE audit_log (
  id          bigserial PRIMARY KEY,
  at          timestamptz NOT NULL DEFAULT now(),
  user_id     uuid REFERENCES users(id),
  action      text NOT NULL CHECK (action IN ('create','update','delete','import','restore','master')),
  entity      text NOT NULL,
  entity_id   text NOT NULL,
  changes     jsonb NOT NULL,          -- [{field, from, to}]
  ip          inet,
  prev_hash   text,                    -- سلسلة تجزئة لكشف العبث
  hash        text NOT NULL
);
REVOKE UPDATE, DELETE ON audit_log FROM PUBLIC;

-- ---------- طرق عرض للتقارير ----------
CREATE VIEW v_project_summary AS
SELECT p.id, p.code, p.name, p.priority, p.budget,
       o.id AS objective_id, o.perspective, d.name AS department,
       coalesce((SELECT sum(amount) FROM project_monthly_budget m WHERE m.project_id = p.id),0) AS monthly_total,
       coalesce((SELECT sum(cost)   FROM project_steps s WHERE s.project_id = p.id),0)         AS steps_cost,
       coalesce(p.manual_progress,
         (SELECT sum(s.pct * s.plan_weeks) / nullif(sum(s.plan_weeks),0) FROM project_steps s WHERE s.project_id = p.id), 0) AS progress
FROM projects p
JOIN objectives o ON o.id = p.objective_id
JOIN departments d ON d.id = p.department_id
WHERE p.deleted_at IS NULL;

-- ---------- أمان على مستوى الصف: محرر الإدارة يرى ويعدّل مشاريع إدارته فقط ----------
ALTER TABLE projects ENABLE ROW LEVEL SECURITY;
CREATE POLICY p_read  ON projects FOR SELECT USING (true);
CREATE POLICY p_write ON projects FOR ALL USING (
  current_setting('app.role') IN ('admin','planner')
  OR (current_setting('app.role') = 'dept_editor' AND department_id = current_setting('app.department_id'))
);
