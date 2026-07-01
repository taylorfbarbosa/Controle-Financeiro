export type Session = {
  user: {
    id: string;
    email: string | null;
    user_metadata: Record<string, unknown>;
  };
};

type AuthEvent = 'SIGNED_IN' | 'SIGNED_OUT' | 'PASSWORD_RECOVERY';
type AuthError = { message: string };
type AuthResult<T = undefined> = { data: T; error: AuthError | null };

const events = new EventTarget();

async function request<T>(method: 'GET' | 'POST', body?: Record<string, unknown>): Promise<T> {
  const response = await fetch('/api/auth', {
    method,
    credentials: 'same-origin',
    headers: body ? { 'Content-Type': 'application/json' } : undefined,
    body: body ? JSON.stringify(body) : undefined,
  });
  const payload = await response.json().catch(() => ({})) as { error?: string } & T;
  if (!response.ok) throw new Error(payload.error || 'Authentication request failed');
  return payload;
}

function emit(event: AuthEvent, session: Session | null) {
  events.dispatchEvent(new CustomEvent('auth', { detail: { event, session } }));
}

async function consumeRecoveryFragment() {
  const hash = new URLSearchParams(window.location.hash.replace(/^#/, ''));
  if (hash.get('type') !== 'recovery') return false;
  const accessToken = hash.get('access_token');
  const refreshToken = hash.get('refresh_token');
  history.replaceState(null, '', `${window.location.pathname}?recovery=1`);
  if (!accessToken || !refreshToken) return false;
  const payload = await request<{ session: Session }>('POST', {
    action: 'establish-recovery', accessToken, refreshToken,
  });
  emit('PASSWORD_RECOVERY', payload.session);
  return true;
}

async function result<T>(operation: () => Promise<T>, fallback: T): Promise<AuthResult<T>> {
  try { return { data: await operation(), error: null }; }
  catch (error) { return { data: fallback, error: { message: error instanceof Error ? error.message : 'Authentication request failed' } }; }
}

export const authClient = {
  auth: {
    async getSession(): Promise<AuthResult<{ session: Session | null }>> {
      return result(async () => {
        await consumeRecoveryFragment();
        const payload = await request<{ session: Session | null }>('GET');
        return { session: payload.session };
      }, { session: null });
    },
    onAuthStateChange(callback: (event: AuthEvent, session: Session | null) => void) {
      const listener = (event: Event) => {
        const detail = (event as CustomEvent<{ event: AuthEvent; session: Session | null }>).detail;
        callback(detail.event, detail.session);
      };
      events.addEventListener('auth', listener);
      return { data: { subscription: { unsubscribe: () => events.removeEventListener('auth', listener) } } };
    },
    signInWithPassword(credentials: { email: string; password: string }) {
      return result<{ session: Session | null }>(async () => {
        const payload = await request<{ session: Session }>('POST', { action: 'login', ...credentials });
        emit('SIGNED_IN', payload.session);
        return { session: payload.session };
      }, { session: null });
    },
    signUp(credentials: { email: string; password: string; fullName: string }) {
      return result(async () => {
        const payload = await request<{ session: Session | null }>('POST', { action: 'signup', ...credentials });
        if (payload.session) emit('SIGNED_IN', payload.session);
        return { session: payload.session };
      }, { session: null });
    },
    resetPasswordForEmail(email: string, _options?: { redirectTo?: string }) {
      return result(async () => {
        await request('POST', { action: 'recovery', email });
        return undefined;
      }, undefined);
    },
    updateUser(update: { password: string }) {
      return result(async () => {
        await request('POST', { action: 'update-password', password: update.password });
        return undefined;
      }, undefined);
    },
    signOut() {
      return result(async () => {
        await request('POST', { action: 'logout' });
        emit('SIGNED_OUT', null);
        return undefined;
      }, undefined);
    },
  },
};
