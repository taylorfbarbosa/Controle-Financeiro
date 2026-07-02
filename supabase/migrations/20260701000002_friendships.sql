-- Friendships: amizades bidirecionais entre usuários.
-- Uma linha representa o pedido de requester_id → receiver_id.
-- Para checar amizade bidirecional, verificar nos dois sentidos.

CREATE TABLE public.friendships (
  id           uuid        PRIMARY KEY DEFAULT gen_random_uuid(),
  requester_id uuid        NOT NULL REFERENCES public.profiles(id) ON DELETE CASCADE,
  receiver_id  uuid        NOT NULL REFERENCES public.profiles(id) ON DELETE CASCADE,
  status       text        NOT NULL DEFAULT 'pending'
               CHECK (status IN ('pending', 'accepted', 'declined')),
  created_at   timestamptz NOT NULL DEFAULT now(),
  updated_at   timestamptz NOT NULL DEFAULT now(),
  CONSTRAINT friendships_no_self CHECK (requester_id <> receiver_id),
  UNIQUE (requester_id, receiver_id)
);

ALTER TABLE public.friendships ENABLE ROW LEVEL SECURITY;

-- Qualquer um dos dois pode ver o registro
CREATE POLICY "friendships_select" ON public.friendships
  FOR SELECT USING (
    auth.uid() = requester_id
    OR auth.uid() = receiver_id
  );

-- Somente o solicitante cria o pedido
CREATE POLICY "friendships_insert" ON public.friendships
  FOR INSERT WITH CHECK (auth.uid() = requester_id);

-- Qualquer um dos dois pode atualizar (aceitar/recusar/cancelar)
CREATE POLICY "friendships_update" ON public.friendships
  FOR UPDATE USING (
    auth.uid() = requester_id
    OR auth.uid() = receiver_id
  );

-- Qualquer um dos dois pode excluir (desfazer amizade)
CREATE POLICY "friendships_delete" ON public.friendships
  FOR DELETE USING (
    auth.uid() = requester_id
    OR auth.uid() = receiver_id
  );

-- Trigger para atualizar updated_at automaticamente
CREATE OR REPLACE FUNCTION public.set_updated_at()
RETURNS TRIGGER LANGUAGE plpgsql AS $$
BEGIN
  NEW.updated_at = now();
  RETURN NEW;
END;
$$;

CREATE TRIGGER friendships_updated_at
  BEFORE UPDATE ON public.friendships
  FOR EACH ROW EXECUTE PROCEDURE public.set_updated_at();

-- Índices
CREATE INDEX friendships_requester_idx ON public.friendships(requester_id);
CREATE INDEX friendships_receiver_idx  ON public.friendships(receiver_id);
CREATE INDEX friendships_status_idx    ON public.friendships(status);
