const base = import.meta.env.BASE_URL;

export default function Slide1Title() {
  return (
    <div
      className="w-screen h-screen overflow-hidden relative"
      style={{
        backgroundColor: "#FAFBFC",
        fontFamily: "'Inter', sans-serif",
        padding: "4vh 4vw",
        boxSizing: "border-box",
        display: "grid",
        gridTemplateColumns: "3fr 2fr",
        gridTemplateRows: "auto 1fr auto",
        gap: "4vh 4vw",
        color: "#1E3A5F",
      }}
    >
      {/* Header */}
      <div
        style={{
          gridColumn: "1 / -1",
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          borderBottom: "1px solid #E2E8F0",
          paddingBottom: "2vh",
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: "1vw" }}>
          <div
            style={{
              width: "2vw",
              height: "2vw",
              backgroundColor: "#0D9488",
              borderRadius: "0.4vw",
            }}
          />
          <div
            style={{
              fontSize: "1.2vw",
              fontWeight: 700,
              letterSpacing: "0.02em",
            }}
          >
            Command Pilot
          </div>
        </div>
        <div
          style={{
            display: "flex",
            gap: "2vw",
            fontSize: "1vw",
            fontWeight: 500,
            color: "#64748B",
          }}
        >
          <div>CP INLINE</div>
          <div>2026</div>
        </div>
      </div>

      {/* Left — hero content */}
      <div
        style={{
          display: "flex",
          flexDirection: "column",
          justifyContent: "center",
          gap: "0",
        }}
      >
        <div
          style={{
            fontSize: "1vw",
            fontWeight: 600,
            color: "#0D9488",
            marginBottom: "1.5vh",
            textTransform: "uppercase",
            letterSpacing: "0.06em",
          }}
        >
          Command Pilot Inline Production Module
        </div>
        <h1
          style={{
            fontSize: "4vw",
            fontWeight: 800,
            margin: "0 0 1.5vh 0",
            lineHeight: 1.05,
            letterSpacing: "-0.02em",
            textWrap: "balance",
          } as React.CSSProperties}
        >
          VCCS (PIPS) Command Pilot™ Inline
        </h1>
        <p
          style={{
            fontSize: "1.5vw",
            fontWeight: 400,
            color: "#475569",
            margin: "0 0 3.5vh 0",
            lineHeight: 1.5,
            maxWidth: "36vw",
          }}
        >
          Fast-track UDI barcode ISO/IEC print quality verification — deployed
          in 6 weeks.
        </p>

        {/* Bullet points */}
        <div style={{ display: "flex", flexDirection: "column", gap: "1.5vh" }}>
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: "1vw",
              background: "#FFFFFF",
              padding: "1.5vh 1.5vw",
              borderRadius: "0.7vw",
              border: "1px solid #E2E8F0",
              boxShadow: "0 0.3vw 1vw rgba(30, 58, 95, 0.05)",
            }}
          >
            <div
              style={{
                width: "0.6vw",
                height: "0.6vw",
                backgroundColor: "#0D9488",
                borderRadius: "50%",
                flexShrink: 0,
              }}
            />
            <div style={{ fontSize: "1.3vw", fontWeight: 500, color: "#1E3A5F" }}>
              Built on proven Command Pilot technology
            </div>
          </div>
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: "1vw",
              background: "#FFFFFF",
              padding: "1.5vh 1.5vw",
              borderRadius: "0.7vw",
              border: "1px solid #E2E8F0",
              boxShadow: "0 0.3vw 1vw rgba(30, 58, 95, 0.05)",
            }}
          >
            <div
              style={{
                width: "0.6vw",
                height: "0.6vw",
                backgroundColor: "#0D9488",
                borderRadius: "50%",
                flexShrink: 0,
              }}
            />
            <div style={{ fontSize: "1.3vw", fontWeight: 500, color: "#1E3A5F" }}>
              Cognex DM-475V-LBL using your existing (adapted) desktop stand
            </div>
          </div>
          <div
            style={{
              display: "flex",
              alignItems: "center",
              gap: "1vw",
              background: "#FFFFFF",
              padding: "1.5vh 1.5vw",
              borderRadius: "0.7vw",
              border: "1px solid #E2E8F0",
              boxShadow: "0 0.3vw 1vw rgba(30, 58, 95, 0.05)",
            }}
          >
            <div
              style={{
                width: "0.6vw",
                height: "0.6vw",
                backgroundColor: "#0D9488",
                borderRadius: "50%",
                flexShrink: 0,
              }}
            />
            <div style={{ fontSize: "1.3vw", fontWeight: 500, color: "#1E3A5F" }}>
              Full-screen operator panel · Indicator pole control · Excel and other reporting
            </div>
          </div>
        </div>
      </div>

      {/* Right — hero image */}
      <div
        style={{
          display: "flex",
          flexDirection: "column",
          justifyContent: "center",
          alignItems: "center",
        }}
      >
        <div
          style={{
            background: "#FFFFFF",
            borderRadius: "1vw",
            border: "1px solid #E2E8F0",
            width: "100%",
            height: "100%",
            overflow: "hidden",
            boxShadow: "0 0.5vw 1.5vw rgba(30, 58, 95, 0.08)",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
          }}
        >
          <img
            src={`${base}hero-scanner.jpg`}
            crossOrigin="anonymous"
            alt="Inline barcode scanner on production conveyor"
            style={{ width: "100%", height: "100%", objectFit: "cover" }}
          />
        </div>
      </div>

      {/* Footer */}
      <div
        style={{
          gridColumn: "1 / -1",
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
          borderTop: "1px solid #E2E8F0",
          paddingTop: "2vh",
          fontSize: "0.9vw",
          color: "#94A3B8",
          fontWeight: 500,
        }}
      >
        <div>VCCS (PIPS)</div>
        <div style={{ display: "flex", gap: "1vw" }}>
          <span>Confidential</span>
          <span>•</span>
          <span>CP Inline — Customer Overview</span>
        </div>
      </div>
    </div>
  );
}
