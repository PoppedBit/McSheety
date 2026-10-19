import type { TableCapabilities } from '@mcsheety/client';
import { createClient } from '@mcsheety/client';
import { useEffect, useMemo, useState } from 'react';
import type { Session } from './auth';
import { DEMO_USERS, demoSession, entraMode, signInWithEntra } from './auth';
import { config, missingEntraConfig } from './config';
import { TableView } from './TableView';

function SignIn({ onDemo }: { onDemo: (userId: string) => void }) {
  const missing = missingEntraConfig();
  if (config.demo) {
    return (
      <div className="card">
        <h2>Demo sign-in</h2>
        <p className="muted">
          Demo mode talks to the local in-memory API (<code>pnpm --filter local-api demo</code>).
          Pick a user to see how roles change what the API returns.
        </p>
        <div className="demo-users">
          {DEMO_USERS.map((u) => (
            <button key={u.id} onClick={() => onDemo(u.id)}>
              {u.label}
            </button>
          ))}
        </div>
      </div>
    );
  }
  return (
    <div className="card">
      <h2>Sign in</h2>
      {missing.length > 0 ? (
        <p className="error">Missing configuration: {missing.join(', ')}. See notes.md.</p>
      ) : (
        <button onClick={() => void signInWithEntra()}>Sign in with Microsoft</button>
      )}
    </div>
  );
}

export function App() {
  const [session, setSession] = useState<Session | null>(null);
  const [restoring, setRestoring] = useState(!config.demo);
  const [caps, setCaps] = useState<TableCapabilities[] | null>(null);
  const [active, setActive] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (config.demo || missingEntraConfig().length > 0) {
      setRestoring(false);
      return;
    }
    entraMode
      .restore()
      .then(setSession)
      .catch(() => setError('Sign-in failed. Try again.'))
      .finally(() => setRestoring(false));
  }, []);

  const db = useMemo(
    () =>
      session
        ? createClient({ apiUrl: config.apiUrl, getToken: (req) => session.getToken(req) })
        : null,
    [session],
  );

  useEffect(() => {
    if (!db) return;
    setCaps(null);
    db.capabilities()
      .then((c) => {
        setCaps(c);
        setActive(c.find((t) => t.table === 'Equipment')?.table ?? c[0]?.table ?? null);
      })
      .catch((e: unknown) =>
        setError(e instanceof Error ? e.message : 'Could not load your permissions.'),
      );
  }, [db]);

  const capability = caps?.find((c) => c.table === active);

  return (
    <div className="app">
      <header>
        <h1>McSheety Equipment</h1>
        {session ? (
          <div className="user">
            <span>{session.label}</span>
            <button className="secondary" onClick={() => void session.signOut()}>
              Sign out
            </button>
          </div>
        ) : null}
      </header>

      <aside className="warning" role="note">
        <strong>About direct Excel access:</strong> McSheety enforces these permissions only for
        requests made through its API. Anyone who can open the workbook in SharePoint or Excel is
        governed by SharePoint permissions instead and can edit any cell — including the permission
        sheets.
      </aside>

      {error ? (
        <div role="alert" className="error">
          {error}
        </div>
      ) : null}

      {restoring ? <p className="muted">Signing in…</p> : null}
      {!restoring && !session ? (
        <SignIn
          onDemo={(id) =>
            setSession(
              demoSession(id, () => {
                setSession(null);
                setCaps(null);
              }),
            )
          }
        />
      ) : null}

      {db && caps ? (
        caps.length === 0 ? (
          <p className="card">
            Your account has no McSheety roles for this workbook. Ask an administrator to assign
            one.
          </p>
        ) : (
          <>
            <nav className="tabs">
              {caps.map((c) => (
                <button
                  key={c.table}
                  className={c.table === active ? 'active' : 'secondary'}
                  onClick={() => setActive(c.table)}
                >
                  {c.table}
                </button>
              ))}
            </nav>
            {active && capability ? (
              <TableView
                key={`${active}-${session?.label}`}
                db={db}
                table={active}
                capability={capability}
              />
            ) : null}
          </>
        )
      ) : null}
    </div>
  );
}
