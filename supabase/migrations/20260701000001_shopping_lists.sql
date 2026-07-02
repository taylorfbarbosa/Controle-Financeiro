-- Shopping lists: cada lista pertence ao criador (user_id = creator_id)
-- Participantes podem ler e atualizar via política separada.

CREATE TABLE public.shopping_lists (
  id               uuid        PRIMARY KEY,
  user_id          uuid        NOT NULL REFERENCES public.profiles(id) ON DELETE CASCADE,
  name             text        NOT NULL CHECK (char_length(name) BETWEEN 1 AND 200),
  date             date,
  note             text        CHECK (char_length(note) <= 2000),
  status           text        NOT NULL DEFAULT 'open'
                               CHECK (status IN ('open', 'finalized', 'cancelled')),
  creator_id       uuid        NOT NULL,
  participant_ids  uuid[]      NOT NULL DEFAULT '{}',
  items            jsonb       NOT NULL DEFAULT '[]',
  created_at       timestamptz NOT NULL DEFAULT now(),
  finalized_at     timestamptz,
  finalized_by_id  uuid,
  cancelled_at     timestamptz
);

ALTER TABLE public.shopping_lists ENABLE ROW LEVEL SECURITY;

-- Criador ou participante pode ler
CREATE POLICY "shopping_lists_select" ON public.shopping_lists
  FOR SELECT USING (
    auth.uid() = user_id
    OR auth.uid() = ANY(participant_ids)
  );

-- Somente criador pode inserir (user_id deve ser o próprio usuário)
CREATE POLICY "shopping_lists_insert" ON public.shopping_lists
  FOR INSERT WITH CHECK (auth.uid() = user_id);

-- Criador ou participante pode atualizar (ex.: adicionar itens)
CREATE POLICY "shopping_lists_update" ON public.shopping_lists
  FOR UPDATE USING (
    auth.uid() = user_id
    OR auth.uid() = ANY(participant_ids)
  );

-- Somente criador pode excluir
CREATE POLICY "shopping_lists_delete" ON public.shopping_lists
  FOR DELETE USING (auth.uid() = user_id);

-- Índice para busca por usuário/participante
CREATE INDEX shopping_lists_user_id_idx   ON public.shopping_lists(user_id);
CREATE INDEX shopping_lists_created_at_idx ON public.shopping_lists(created_at DESC);
