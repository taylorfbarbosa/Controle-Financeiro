# Segurança e implantação

## Arquitetura

- O navegador não recebe tokens do Supabase. Sessões usam cookies `HttpOnly`, `Secure` em produção e `SameSite=Strict`.
- Login e acesso aos dados passam pelas funções `/api/auth` e `/api/data`.
- A Service Role é usada somente no servidor para o contador persistente de rate limit e revogação de sessão.
- Toda entrada enviada às APIs é validada com Zod e toda consulta de usuário continua protegida por RLS.

## Configuração obrigatória

1. Aplicar as migrations de `supabase/migrations` no projeto Supabase.
2. Configurar na Vercel, como variáveis protegidas:
   - `SUPABASE_URL`
   - `SUPABASE_ANON_KEY`
   - `SUPABASE_SERVICE_ROLE_KEY`
3. Nunca criar uma variável `VITE_SUPABASE_SERVICE_ROLE_KEY`. Todo valor com prefixo `VITE_` pode ser exposto ao cliente.
4. Implantar somente após `npm run lint`, `npm run typecheck`, `npm run build` e `npm audit --omit=dev` passarem.

Sem a Service Role em produção, as APIs falham de forma segura para não operar sem rate limit persistente.
