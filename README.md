# Cracha Fast

SaaS para cadastro de pessoas, modelos e geração de carteirinhas em lote.

Stack: React, TypeScript, Vite, Supabase, XLSX/CSV, Word export.

MVP:
- Dashboard
- Importação XLSX, XLS e CSV
- Cadastro e edição de pessoas
- Modelo inicial Espaço Confinado
- Editor rápido de modelo
- Geração em lote
- 12 crachás por A4
- Impressão/PDF
- Word editável
- Histórico local
- Schema Supabase com multiempresa e RLS

Rodar:
npm install
npm run dev

Supabase:
copie .env.example para .env e preencha VITE_SUPABASE_URL e VITE_SUPABASE_ANON_KEY.
Execute supabase/schema.sql no SQL Editor.

Próximas evoluções: autenticação Supabase, armazenamento multiempresa, logos por empresa, editor visual completo, assinatura e cobrança.