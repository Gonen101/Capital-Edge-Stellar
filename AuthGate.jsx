// ============================================================================
// AuthGate.jsx
//
// Sits in front of the whole app. Logged out -> shows the animated landing
// page with login/signup. Logged in -> loads that user's own company row
// from Supabase and renders <App/>, passing it the company data plus a
// save function.
//
// The hero animation (figures, chart, trophy) is the same one built and
// approved as a standalone preview - ported directly into this component
// rather than rebuilt, so what you saw in the preview is exactly what ships.
// The only thing that changed is the form itself: the decorative sign-in
// box from the preview is replaced here with the real, working Supabase
// auth logic from the original AuthGate.
// ============================================================================
import { useState, useEffect, useCallback, useRef } from 'react';
import { supabase } from './supabaseClient';
import App from './App'; // ces.jsx, after the data-layer patch in DEPLOYMENT-GUIDE.md

export default function AuthGate() {
  const [session, setSession] = useState(undefined); // undefined = still checking, null = logged out
  const [company, setCompany] = useState(null);       // { id, name, data }
  const [loadingCompany, setLoadingCompany] = useState(false);
  const [error, setError] = useState('');

  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => setSession(session));
    const { data: listener } = supabase.auth.onAuthStateChange((_event, session) => setSession(session));
    return () => listener.subscription.unsubscribe();
  }, []);

  useEffect(() => {
    if (!session) { setCompany(null); return; }
    setLoadingCompany(true);
    supabase
      .from('companies')
      .select('id, name, data')
      .eq('owner_id', session.user.id)
      .single()
      .then(({ data, error }) => {
        if (error) setError('Could not load your company: ' + error.message);
        else setCompany(data);
        setLoadingCompany(false);
      });
  }, [session]);

  const saveTimer = useRef(null);
  const saveCompanyData = useCallback((newData) => {
    if (!company) return;
    setCompany((c) => (c ? { ...c, data: newData } : c));
    clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(async () => {
      const { error } = await supabase.from('companies').update({ data: newData }).eq('id', company.id);
      if (error) console.error('Save failed:', error.message);
    }, 400);
  }, [company]);

  if (session === undefined) return <CenteredMessage>Loading...</CenteredMessage>;
  if (!session) return <LandingPage onError={setError} error={error} />;
  if (loadingCompany || !company) return <CenteredMessage>Loading your company...</CenteredMessage>;

  return <App companyData={company.data} onCompanyDataChange={saveCompanyData} onSignOut={() => supabase.auth.signOut()} />;
}

function CenteredMessage({ children }) {
  return (
    <div style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', fontFamily: 'Inter, sans-serif', color: '#475569' }}>
      {children}
    </div>
  );
}

const LANDING_CSS = `
  .cel-chip { position: absolute; background: rgba(37,99,235,0.16); border: 1px solid rgba(37,99,235,0.35); color: #C6D8FC; font-size: 13px; font-weight: 500; padding: 6px 12px; border-radius: 20px; animation: celChipScatter 12s ease-in-out infinite; white-space: nowrap; }
  .cel-chip b { color: #F2F5FA; font-weight: 700; }
  .cel-bar { transform-origin: bottom; animation: celBarGrow 12s ease-out infinite; }
  .cel-trend-line { stroke-dasharray: 420; stroke-dashoffset: 420; animation: celLineDraw 12s ease-out infinite; }
  .cel-trophy { opacity: 0; animation: celTrophyIn 12s ease-out infinite; transform-origin: 315px 260px; }
  .cel-trophy-shine { opacity: 0; animation: celTrophyGlow 12s ease-out infinite; transform-origin: 315px 245px; }
  .cel-climber { opacity: 0; animation: celClimb 12s ease-in-out infinite; }
  .cel-page-flap { animation: celFlutter 0.5s ease-in-out infinite alternate; transform-origin: 0px -4px; }
  @keyframes celChipScatter {
    0%, 4% { opacity: 1; transform: translateY(0) scale(1); }
    30% { opacity: 1; transform: translateY(-30px) scale(1); }
    38%, 100% { opacity: 0; transform: translateY(-50px) scale(0.9); }
  }
  @keyframes celBarGrow {
    0%, 5% { transform: scaleY(0); }
    40%, 90% { transform: scaleY(1); }
    97%, 100% { transform: scaleY(0); opacity: 0.4; }
  }
  @keyframes celLineDraw {
    0%, 36% { stroke-dashoffset: 420; opacity: 0; }
    40% { opacity: 1; }
    54%, 92% { stroke-dashoffset: 0; opacity: 1; }
    98%, 100% { opacity: 0; }
  }
  @keyframes celTrophyIn {
    0%, 46% { opacity: 0; transform: scale(0.7); }
    54%, 96% { opacity: 1; transform: scale(1); }
    100% { opacity: 0; transform: scale(0.85); }
  }
  @keyframes celTrophyGlow {
    0%, 90% { opacity: 0; transform: scale(0.8); }
    94%, 98% { opacity: 1; transform: scale(1.2); }
    100% { opacity: 0; transform: scale(1); }
  }
  @keyframes celClimb {
    0%, 55% { opacity: 0; transform: translate(30px, 262px); }
    59% { opacity: 1; transform: translate(30px, 262px); }
    65% { transform: translate(65px, 202px); }
    71% { transform: translate(115px, 172px); }
    77% { transform: translate(165px, 142px); }
    83% { transform: translate(215px, 112px); }
    89% { transform: translate(265px, 82px); }
    94%, 98% { opacity: 1; transform: translate(315px, 52px); }
    100% { opacity: 0; transform: translate(315px, 52px); }
  }
  @keyframes celFlutter {
    0% { transform: skewY(-6deg); }
    100% { transform: skewY(6deg); }
  }
`;

// The bars/trend-line/trophy/climber chart, reused twice - once sharp on
// the hero side, once heavily blurred behind the sign-in form. Both copies
// share the same CSS classes and keyframes above, so they stay perfectly in
// sync with no extra wiring.
function ChartScene({ gradId }) {
  return (
    <svg viewBox="0 0 380 280" style={{ position: 'absolute', inset: 0, width: '100%', height: '100%' }}>
      <line x1="20" y1="270" x2="360" y2="270" stroke="rgba(139,151,184,0.4)" strokeWidth="1" />
      <rect className="cel-bar" x="48" y="210" width="34" height="60" rx="4" fill={`url(#${gradId})`} />
      <rect className="cel-bar" x="98" y="180" width="34" height="90" rx="4" fill={`url(#${gradId})`} style={{ animationDelay: '0.08s' }} />
      <rect className="cel-bar" x="148" y="150" width="34" height="120" rx="4" fill={`url(#${gradId})`} style={{ animationDelay: '0.16s' }} />
      <rect className="cel-bar" x="198" y="120" width="34" height="150" rx="4" fill={`url(#${gradId})`} style={{ animationDelay: '0.24s' }} />
      <rect className="cel-bar" x="248" y="90" width="34" height="180" rx="4" fill={`url(#${gradId})`} style={{ animationDelay: '0.32s' }} />
      <rect className="cel-bar" x="298" y="60" width="34" height="210" rx="4" fill={`url(#${gradId})`} style={{ animationDelay: '0.4s' }} />
      <polyline className="cel-trend-line" points="65,204 115,174 165,144 215,114 265,84 315,54" fill="none" stroke="#4FA6F7" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" />
      <g className="cel-trophy">
        <ellipse cx="315" cy="60" rx="8" ry="3" fill="#8A6A1E" />
        <rect x="309" y="52" width="12" height="9" fill="#D9A62E" />
        <path d="M300 28 Q300 46 315 46 Q330 46 330 28 L330 22 L300 22 Z" fill="#F0C244" />
        <path d="M300 24 Q290 24 290 32 Q290 40 300 39" stroke="#F0C244" strokeWidth="3" fill="none" />
        <path d="M330 24 Q340 24 340 32 Q340 40 330 39" stroke="#F0C244" strokeWidth="3" fill="none" />
      </g>
      <g className="cel-trophy-shine">
        <line x1="315" y1="12" x2="315" y2="2" stroke="#FBE7A6" strokeWidth="2" strokeLinecap="round" />
        <line x1="295" y1="20" x2="287" y2="14" stroke="#FBE7A6" strokeWidth="2" strokeLinecap="round" />
        <line x1="335" y1="20" x2="343" y2="14" stroke="#FBE7A6" strokeWidth="2" strokeLinecap="round" />
      </g>
      <g className="cel-climber">
        <ellipse cx="0" cy="9" rx="13" ry="3" fill="rgba(0,0,0,0.25)" />
        <path d="M-7 6 L-6 -8 Q-6 -14 0 -14 Q6 -14 6 -8 L7 6 Z" fill="#1E5FD9" />
        <circle cx="0" cy="-22" r="7" fill="#F2C9A0" />
        <g className="cel-page-flap"><path d="M-10 -4 L0 -1 L0 7 L-10 4 Z" fill="#F7FAFF" /></g>
        <path d="M0 -1 L10 -4 L10 4 L0 7 Z" fill="#EDF1F8" />
        <path d="M-5 -10 L-9 -4" stroke="#F2C9A0" strokeWidth="4" strokeLinecap="round" />
        <path d="M5 -10 L9 -4" stroke="#F2C9A0" strokeWidth="4" strokeLinecap="round" />
      </g>
      <defs>
        <linearGradient id={gradId} x1="0" y1="1" x2="0" y2="0">
          <stop offset="0%" stopColor="#2563EB" />
          <stop offset="100%" stopColor="#4F8AFA" />
        </linearGradient>
      </defs>
    </svg>
  );
}

function BrandMark() {
  return (
    <svg width="52" height="52" viewBox="0 0 100 100" style={{ display: 'block', margin: '0 auto' }}>
      <defs>
        <linearGradient id="celCrescent" x1="0%" y1="0%" x2="100%" y2="100%"><stop offset="0%" stopColor="#0B1F4D" /><stop offset="100%" stopColor="#16336E" /></linearGradient>
        <linearGradient id="celLogoArrow" x1="0%" y1="100%" x2="100%" y2="0%"><stop offset="0%" stopColor="#1E5FD9" /><stop offset="100%" stopColor="#5CB0FA" /></linearGradient>
      </defs>
      <circle cx="50" cy="50" r="46" fill="url(#celCrescent)" />
      <circle cx="61" cy="40" r="39" fill="#0B1730" />
      <rect x="27" y="62" width="7" height="12" rx="1.3" fill="url(#celLogoArrow)" />
      <rect x="38" y="54" width="7" height="20" rx="1.3" fill="url(#celLogoArrow)" />
      <rect x="49" y="46" width="7" height="28" rx="1.3" fill="url(#celLogoArrow)" />
      <rect x="60" y="38" width="7" height="36" rx="1.3" fill="url(#celLogoArrow)" />
      <path d="M29 68 Q 54 54 76 30" stroke="url(#celLogoArrow)" strokeWidth="4.5" fill="none" strokeLinecap="round" />
      <path d="M67 21 L80 25 L76 38 Z" fill="url(#celLogoArrow)" />
    </svg>
  );
}

function LandingPage({ error, onError }) {
  return (
    <div style={{ display: 'flex', minHeight: '100vh', fontFamily: 'Inter, sans-serif' }}>
      <style>{LANDING_CSS}</style>
      <div style={{ flex: 1.15, position: 'relative', background: '#0B1730', overflow: 'hidden', display: 'flex', flexDirection: 'column', alignItems: 'center', padding: '60px 40px 40px' }}>
        <div style={{ position: 'absolute', inset: 0, backgroundImage: 'radial-gradient(rgba(255,255,255,0.05) 1px, transparent 1px)', backgroundSize: '26px 26px', opacity: 0.5 }} />
        <div style={{ position: 'relative', zIndex: 2, textAlign: 'center', maxWidth: 400 }}>
          <div style={{ fontSize: 28, fontWeight: 700, color: '#F2F5FA', letterSpacing: '-0.02em', lineHeight: 1.3 }}>Every figure, reconciled.<br />Every close, effortless.</div>
        </div>
        <div style={{ position: 'relative', width: 380, height: 280, margin: '36px auto 0' }}>
          <span className="cel-chip" style={{ left: '4%', bottom: 8 }}><b>4.8M</b> revenue</span>
          <span className="cel-chip" style={{ left: '58%', bottom: 2, animationDelay: '0.3s' }}><b>+24%</b> YoY</span>
          <span className="cel-chip" style={{ left: '28%', bottom: 18, animationDelay: '0.6s' }}><b>2.7M</b> net profit</span>
          <span className="cel-chip" style={{ left: '76%', bottom: 10, animationDelay: '0.15s' }}><b>128</b> invoices</span>
          <ChartScene gradId="celBarGradHero" />
        </div>
        <div style={{ position: 'relative', zIndex: 2, marginTop: 'auto', paddingTop: 26, textAlign: 'center' }}>
          <BrandMark />
          <div style={{ fontFamily: "'Playfair Display', serif", fontWeight: 800, fontSize: 20, color: '#F2F5FA', letterSpacing: '-0.01em', marginTop: 8 }}>CapitalEdge Stellar</div>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 10, marginTop: 2 }}>
            <span style={{ width: 20, height: 1, background: '#4A5578' }} />
            <span style={{ fontFamily: "'Playfair Display', serif", fontSize: 11, letterSpacing: '0.15em', color: '#8B97B8' }}>LTD</span>
            <span style={{ width: 20, height: 1, background: '#4A5578' }} />
          </div>
        </div>
      </div>

      <div style={{ flex: 1, position: 'relative', background: '#0B1730', overflow: 'hidden', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '40px 24px' }}>
        <div style={{ position: 'absolute', inset: -40, filter: 'blur(30px)', opacity: 0.6, transform: 'scale(1.35)' }}>
          <div style={{ position: 'relative', width: 380, height: 280, margin: '90px auto 0' }}>
            <ChartScene gradId="celBarGradBg" />
          </div>
        </div>
        <div style={{ position: 'absolute', inset: 0, background: 'radial-gradient(ellipse at center, rgba(11,23,48,0.55) 0%, rgba(11,23,48,0.88) 70%)' }} />
        <AuthForm error={error} onError={onError} />
      </div>
    </div>
  );
}

function AuthForm({ error, onError }) {
  const [mode, setMode] = useState('login'); // 'login' | 'signup'
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [companyName, setCompanyName] = useState('');
  const [submitting, setSubmitting] = useState(false);
  const [notice, setNotice] = useState('');

  const submit = async (e) => {
    e.preventDefault();
    onError('');
    setNotice('');
    setSubmitting(true);
    try {
      if (mode === 'signup') {
        const { error } = await supabase.auth.signUp({
          email,
          password,
          options: { data: { company_name: companyName || 'My Company' } },
        });
        if (error) throw error;
        setNotice('Check your email to confirm your account, then log in below.');
        setMode('login');
      } else {
        const { error } = await supabase.auth.signInWithPassword({ email, password });
        if (error) throw error;
      }
    } catch (err) {
      onError(err.message);
    } finally {
      setSubmitting(false);
    }
  };

  const inputStyle = { width: '100%', padding: '11px 13px', borderRadius: 8, border: '1px solid #D0D5DD', fontSize: 14.5, fontFamily: 'Inter, sans-serif', boxSizing: 'border-box' };
  const labelStyle = { fontSize: 13, fontWeight: 600, color: '#344054', marginBottom: 6, display: 'block' };

  return (
    <form onSubmit={submit} style={{ position: 'relative', zIndex: 2, width: 340, background: '#ffffff', borderRadius: 16, padding: '36px 32px', boxShadow: '0 20px 60px rgba(0,0,0,0.35), 0 0 0 1px rgba(255,255,255,0.06)' }}>
      <div style={{ fontSize: 22, fontWeight: 700, color: '#101828', letterSpacing: '-0.02em' }}>{mode === 'login' ? 'Sign in' : 'Create your company'}</div>
      <div style={{ fontSize: 13.5, color: '#667085', marginTop: 4 }}>
        {mode === 'login' ? 'Welcome back.' : 'Starts completely empty - nothing has been pre-filled for you.'}
      </div>

      {mode === 'signup' && (
        <div style={{ marginTop: 20 }}>
          <label style={labelStyle}>Company name</label>
          <input style={inputStyle} value={companyName} onChange={(e) => setCompanyName(e.target.value)} placeholder="Acme Ltd" />
        </div>
      )}
      <div style={{ marginTop: 20 }}>
        <label style={labelStyle}>Email</label>
        <input style={inputStyle} type="email" required value={email} onChange={(e) => setEmail(e.target.value)} placeholder="you@company.com" />
      </div>
      <div style={{ marginTop: 14 }}>
        <label style={labelStyle}>Password</label>
        <input style={inputStyle} type="password" required minLength={6} value={password} onChange={(e) => setPassword(e.target.value)} placeholder="At least 6 characters" />
      </div>

      {error && <div style={{ marginTop: 14, fontSize: 12.5, color: '#B42318', background: '#FEF3F2', border: '1px solid #FECDCA', borderRadius: 8, padding: '8px 10px' }}>{error}</div>}
      {notice && <div style={{ marginTop: 14, fontSize: 12.5, color: '#175CD3', background: '#EFF8FF', border: '1px solid #B2DDFF', borderRadius: 8, padding: '8px 10px' }}>{notice}</div>}

      <button type="submit" disabled={submitting} style={{ width: '100%', marginTop: 18, padding: '12px 0', borderRadius: 8, border: 'none', background: '#2563EB', color: '#fff', fontWeight: 600, fontSize: 15, cursor: submitting ? 'default' : 'pointer', opacity: submitting ? 0.7 : 1 }}>
        {submitting ? 'Please wait...' : mode === 'login' ? 'Sign in' : 'Create company'}
      </button>

      <div style={{ marginTop: 16, textAlign: 'center', fontSize: 13.5, color: '#667085' }}>
        {mode === 'login' ? (
          <>New here? <a href="#" onClick={(e) => { e.preventDefault(); setMode('signup'); onError(''); }} style={{ color: '#2563EB', fontWeight: 600, textDecoration: 'none' }}>Create a company</a></>
        ) : (
          <>Already have an account? <a href="#" onClick={(e) => { e.preventDefault(); setMode('login'); onError(''); }} style={{ color: '#2563EB', fontWeight: 600, textDecoration: 'none' }}>Sign in</a></>
        )}
      </div>
    </form>
  );
}
