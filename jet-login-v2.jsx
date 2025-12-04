import { useState, useEffect } from 'react';

export default function LoginPage() {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [focused, setFocused] = useState(null);
  const [isLoading, setIsLoading] = useState(false);
  const [mousePos, setMousePos] = useState({ x: 50, y: 50 });

  useEffect(() => {
    const handleMouseMove = (e) => {
      setMousePos({
        x: (e.clientX / window.innerWidth) * 100,
        y: (e.clientY / window.innerHeight) * 100,
      });
    };
    window.addEventListener('mousemove', handleMouseMove);
    return () => window.removeEventListener('mousemove', handleMouseMove);
  }, []);

  const handleSubmit = (e) => {
    e.preventDefault();
    setIsLoading(true);
    setTimeout(() => setIsLoading(false), 2000);
  };

  // JET Logo - Accurate recreation with italic style and swoosh
  const JetLogo = () => (
    <svg width="280" height="100" viewBox="0 0 280 100" fill="none">
      {/* J with teardrop loop */}
      <path 
        d="M45 15 
           L75 15 
           L75 20 
           L55 20 
           L55 55 
           Q55 75 40 80 
           Q25 85 15 75 
           Q8 68 12 55 
           Q16 42 30 38 
           Q38 36 45 40
           L45 15 Z"
        fill="rgba(230, 235, 245, 0.95)"
      />
      {/* J inner curve */}
      <path 
        d="M35 50 
           Q28 52 25 58 
           Q22 65 28 70 
           Q34 75 42 72 
           Q48 69 48 60 
           L48 50 
           Q42 48 35 50 Z"
        fill="#1a1f35"
      />
      
      {/* E */}
      <path 
        d="M85 15 L130 15 L130 22 L95 22 L95 45 L125 45 L125 52 L95 52 L95 73 L132 73 L132 80 L85 80 Z"
        fill="rgba(230, 235, 245, 0.95)"
        style={{ transform: 'skewX(-12deg)', transformOrigin: '107px 47px' }}
      />
      
      {/* T */}
      <path 
        d="M135 15 L190 15 L190 22 L168 22 L168 80 L157 80 L157 22 L135 22 Z"
        fill="rgba(230, 235, 245, 0.95)"
        style={{ transform: 'skewX(-12deg)', transformOrigin: '162px 47px' }}
      />
      
      {/* Diagonal swoosh line through letters */}
      <path 
        d="M20 72 Q60 55 120 35 Q180 15 240 8"
        stroke="rgba(230, 235, 245, 0.95)"
        strokeWidth="2.5"
        strokeLinecap="round"
        fill="none"
      />
      
      {/* Small aircraft/point at end of swoosh */}
      <path 
        d="M238 8 L248 5 L244 12 Z"
        fill="rgba(230, 235, 245, 0.95)"
      />
    </svg>
  );

  return (
    <div style={{
      minHeight: '100vh',
      background: '#1a1f35',
      display: 'flex',
      flexDirection: 'column',
      alignItems: 'center',
      justifyContent: 'center',
      fontFamily: '"DM Sans", -apple-system, sans-serif',
      position: 'relative',
      overflow: 'hidden',
    }}>
      {/* Google Fonts */}
      <link href="https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,400;0,500;1,400&family=DM+Sans:wght@400;500&display=swap" rel="stylesheet" />

      {/* Animated Silk Gradient Background */}
      <div style={{
        position: 'absolute',
        inset: 0,
        background: `
          radial-gradient(ellipse 100% 60% at ${mousePos.x}% ${mousePos.y}%, 
            rgba(99, 115, 155, 0.35) 0%, 
            transparent 50%),
          radial-gradient(ellipse 70% 90% at ${100 - mousePos.x}% ${100 - mousePos.y}%, 
            rgba(71, 85, 125, 0.25) 0%, 
            transparent 50%),
          radial-gradient(ellipse 50% 50% at 80% 20%, 
            rgba(90, 105, 145, 0.2) 0%, 
            transparent 40%),
          linear-gradient(160deg, 
            #171c30 0%, 
            #1e2440 20%, 
            #252d48 40%, 
            #2a3352 60%, 
            #1e2440 80%, 
            #171c30 100%)
        `,
        transition: 'background 1s cubic-bezier(0.4, 0, 0.2, 1)',
      }} />

      {/* Silk Texture */}
      <div style={{
        position: 'absolute',
        inset: 0,
        background: `
          repeating-linear-gradient(
            120deg,
            transparent 0px,
            transparent 2px,
            rgba(255, 255, 255, 0.008) 2px,
            rgba(255, 255, 255, 0.008) 4px
          )
        `,
      }} />

      {/* Flowing Light */}
      <div style={{
        position: 'absolute',
        inset: 0,
        background: `
          linear-gradient(
            ${140 + (mousePos.x - 50) * 0.2}deg,
            transparent 0%,
            rgba(180, 190, 215, 0.04) 25%,
            rgba(200, 210, 235, 0.07) 50%,
            rgba(180, 190, 215, 0.04) 75%,
            transparent 100%
          )
        `,
        transition: 'background 1.5s ease',
      }} />

      {/* Main Content */}
      <div style={{
        position: 'relative',
        zIndex: 1,
        width: '100%',
        maxWidth: '480px',
        padding: '0 40px',
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
      }}>
        
        {/* Logo & Subtitle Section */}
        <div style={{
          width: '100%',
          marginBottom: '56px',
          position: 'relative',
        }}>
          {/* Logo - Left aligned, larger */}
          <div style={{
            display: 'flex',
            justifyContent: 'flex-start',
            marginLeft: '-20px',
          }}>
            <JetLogo />
          </div>
          
          {/* Subtitle - Right side, below logo */}
          <p style={{
            position: 'absolute',
            right: '0',
            bottom: '-8px',
            fontSize: '13px',
            color: 'rgba(180, 190, 215, 0.6)',
            fontFamily: '"Cormorant Garamond", serif',
            fontStyle: 'italic',
            letterSpacing: '0.03em',
            margin: 0,
          }}>
            Precision in Motion
          </p>
        </div>

        {/* Form */}
        <form onSubmit={handleSubmit} style={{ width: '100%' }}>
          {/* Email Input */}
          <div style={{ marginBottom: '20px' }}>
            <input
              type="email"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              onFocus={() => setFocused('email')}
              onBlur={() => setFocused(null)}
              placeholder="Email"
              style={{
                width: '100%',
                padding: '18px 24px',
                fontSize: '15px',
                fontFamily: '"DM Sans", sans-serif',
                background: focused === 'email' 
                  ? 'rgba(255, 255, 255, 0.12)' 
                  : 'rgba(255, 255, 255, 0.08)',
                border: focused === 'email'
                  ? '1px solid rgba(200, 210, 235, 0.3)'
                  : '1px solid rgba(255, 255, 255, 0.1)',
                borderRadius: '8px',
                color: 'rgba(235, 240, 250, 0.95)',
                outline: 'none',
                transition: 'all 0.3s ease',
                boxSizing: 'border-box',
              }}
            />
          </div>

          {/* Password Input */}
          <div style={{ marginBottom: '28px' }}>
            <input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              onFocus={() => setFocused('password')}
              onBlur={() => setFocused(null)}
              placeholder="Password"
              style={{
                width: '100%',
                padding: '18px 24px',
                fontSize: '15px',
                fontFamily: '"DM Sans", sans-serif',
                background: focused === 'password' 
                  ? 'rgba(255, 255, 255, 0.12)' 
                  : 'rgba(255, 255, 255, 0.08)',
                border: focused === 'password'
                  ? '1px solid rgba(200, 210, 235, 0.3)'
                  : '1px solid rgba(255, 255, 255, 0.1)',
                borderRadius: '8px',
                color: 'rgba(235, 240, 250, 0.95)',
                outline: 'none',
                transition: 'all 0.3s ease',
                boxSizing: 'border-box',
              }}
            />
          </div>

          {/* Login Button */}
          <button
            type="submit"
            disabled={isLoading}
            style={{
              width: '100%',
              padding: '18px',
              fontSize: '14px',
              fontWeight: '500',
              fontFamily: '"DM Sans", sans-serif',
              letterSpacing: '0.08em',
              textTransform: 'uppercase',
              background: isLoading 
                ? 'rgba(180, 190, 215, 0.3)'
                : 'rgba(230, 235, 250, 0.95)',
              border: 'none',
              borderRadius: '8px',
              color: '#1a1f35',
              cursor: isLoading ? 'wait' : 'pointer',
              transition: 'all 0.3s ease',
            }}
            onMouseOver={(e) => {
              if (!isLoading) {
                e.target.style.background = 'rgba(255, 255, 255, 1)';
                e.target.style.transform = 'translateY(-2px)';
                e.target.style.boxShadow = '0 8px 32px rgba(0, 0, 0, 0.2)';
              }
            }}
            onMouseOut={(e) => {
              if (!isLoading) {
                e.target.style.background = 'rgba(230, 235, 250, 0.95)';
                e.target.style.transform = 'translateY(0)';
                e.target.style.boxShadow = 'none';
              }
            }}
          >
            {isLoading ? (
              <span style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '10px' }}>
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" style={{ animation: 'spin 1s linear infinite' }}>
                  <circle cx="12" cy="12" r="10" strokeOpacity="0.3"/>
                  <path d="M12 2a10 10 0 0 1 10 10"/>
                </svg>
                Processing
              </span>
            ) : 'Sign In'}
          </button>
        </form>

        {/* Forgot Password */}
        <a href="#" style={{
          marginTop: '24px',
          fontSize: '13px',
          color: 'rgba(180, 190, 215, 0.5)',
          textDecoration: 'none',
          transition: 'color 0.2s',
        }} onMouseOver={(e) => e.target.style.color = 'rgba(200, 210, 235, 0.8)'}
           onMouseOut={(e) => e.target.style.color = 'rgba(180, 190, 215, 0.5)'}>
          Forgot password?
        </a>
      </div>

      {/* Bottom Footer */}
      <div style={{
        position: 'absolute',
        bottom: '40px',
        left: 0,
        right: 0,
        display: 'flex',
        justifyContent: 'center',
        alignItems: 'center',
        gap: '20px',
      }}>
        <div style={{
          width: '60px',
          height: '1px',
          background: 'linear-gradient(90deg, transparent, rgba(180, 190, 215, 0.3))',
        }} />
        <span style={{
          fontSize: '10px',
          color: 'rgba(180, 190, 215, 0.4)',
          letterSpacing: '0.25em',
          fontFamily: '"DM Sans", sans-serif',
          textTransform: 'uppercase',
        }}>Enterprise Platform</span>
        <div style={{
          width: '60px',
          height: '1px',
          background: 'linear-gradient(90deg, rgba(180, 190, 215, 0.3), transparent)',
        }} />
      </div>

      {/* CSS */}
      <style>{`
        @keyframes spin {
          from { transform: rotate(0deg); }
          to { transform: rotate(360deg); }
        }
        
        input::placeholder {
          color: rgba(180, 190, 215, 0.4);
        }
        
        * {
          -webkit-font-smoothing: antialiased;
          -moz-osx-font-smoothing: grayscale;
        }
      `}</style>
    </div>
  );
}
