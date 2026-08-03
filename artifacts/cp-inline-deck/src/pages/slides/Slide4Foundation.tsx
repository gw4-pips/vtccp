export default function Slide4Foundation() {
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
          <div>OUR FOUNDATION</div>
          <div>2026</div>
        </div>
      </div>

      {/* Left — bullet points */}
      <div
        style={{
          display: "flex",
          flexDirection: "column",
          justifyContent: "center",
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
          Built on a solid foundation
        </div>
        <h2
          style={{
            fontSize: "3.5vw",
            fontWeight: 800,
            margin: "0 0 3vh 0",
            lineHeight: 1.1,
            letterSpacing: "-0.02em",
            textWrap: "balance",
          } as React.CSSProperties}
        >
          Built on a solid foundation
        </h2>

        <div style={{ display: "flex", flexDirection: "column", gap: "1.8vh" }}>
          {/* Item 1 */}
          <div
            style={{
              display: "flex",
              gap: "1.5vw",
              alignItems: "flex-start",
              background: "#FFFFFF",
              padding: "1.8vh 1.8vw",
              borderRadius: "0.7vw",
              border: "1px solid #E2E8F0",
              boxShadow: "0 0.2vw 0.8vw rgba(30, 58, 95, 0.04)",
            }}
          >
            <div
              style={{
                fontSize: "1vw",
                fontWeight: 700,
                color: "#0D9488",
                backgroundColor: "rgba(13, 148, 136, 0.1)",
                width: "2.4vw",
                height: "2.4vw",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                borderRadius: "50%",
                flexShrink: 0,
              }}
            >
              1
            </div>
            <div
              style={{
                fontSize: "1.15vw",
                color: "#475569",
                lineHeight: 1.45,
              }}
            >
              Command Pilot's DMCC client, GS1 decoder, and Excel engine have been under development as a desktop tool for months and used internally at PIPS
            </div>
          </div>

          {/* Item 2 */}
          <div
            style={{
              display: "flex",
              gap: "1.5vw",
              alignItems: "flex-start",
              background: "#FFFFFF",
              padding: "1.8vh 1.8vw",
              borderRadius: "0.7vw",
              border: "1px solid #E2E8F0",
              boxShadow: "0 0.2vw 0.8vw rgba(30, 58, 95, 0.04)",
            }}
          >
            <div
              style={{
                fontSize: "1vw",
                fontWeight: 700,
                color: "#0D9488",
                backgroundColor: "rgba(13, 148, 136, 0.1)",
                width: "2.4vw",
                height: "2.4vw",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                borderRadius: "50%",
                flexShrink: 0,
              }}
            >
              2
            </div>
            <div
              style={{
                fontSize: "1.15vw",
                color: "#475569",
                lineHeight: 1.45,
              }}
            >
              No new scanner protocol work, no new parsing logic — CP Inline reuses what already ships
            </div>
          </div>

          {/* Item 3 */}
          <div
            style={{
              display: "flex",
              gap: "1.5vw",
              alignItems: "flex-start",
              background: "#FFFFFF",
              padding: "1.8vh 1.8vw",
              borderRadius: "0.7vw",
              border: "1px solid #E2E8F0",
              boxShadow: "0 0.2vw 0.8vw rgba(30, 58, 95, 0.04)",
            }}
          >
            <div
              style={{
                fontSize: "1vw",
                fontWeight: 700,
                color: "#0D9488",
                backgroundColor: "rgba(13, 148, 136, 0.1)",
                width: "2.4vw",
                height: "2.4vw",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                borderRadius: "50%",
                flexShrink: 0,
              }}
            >
              4
            </div>
            <div
              style={{
                fontSize: "1.15vw",
                color: "#475569",
                lineHeight: 1.45,
              }}
            >
              Windows Remote Desktop for remote support — not blocked by CP Inline
            </div>
          </div>

          {/* Item 4 */}
          <div
            style={{
              display: "flex",
              gap: "1.5vw",
              alignItems: "flex-start",
              background: "#FFFFFF",
              padding: "1.8vh 1.8vw",
              borderRadius: "0.7vw",
              border: "1px solid #E2E8F0",
              boxShadow: "0 0.2vw 0.8vw rgba(30, 58, 95, 0.04)",
            }}
          >
            <div
              style={{
                fontSize: "1vw",
                fontWeight: 700,
                color: "#0D9488",
                backgroundColor: "rgba(13, 148, 136, 0.1)",
                width: "2.4vw",
                height: "2.4vw",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                borderRadius: "50%",
                flexShrink: 0,
              }}
            >
              5
            </div>
            <div
              style={{
                fontSize: "1.15vw",
                color: "#475569",
                lineHeight: 1.45,
              }}
            >
              Audit trail baseline included: operator login, config change log, read-only session records
            </div>
          </div>
        </div>
      </div>

      {/* Right — ~70% stat card */}
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
            padding: "4vh 3vw",
            borderRadius: "1vw",
            border: "1px solid #E2E8F0",
            width: "100%",
            height: "100%",
            display: "flex",
            flexDirection: "column",
            justifyContent: "center",
            alignItems: "center",
            textAlign: "center",
            boxSizing: "border-box",
            boxShadow: "0 0.5vw 1.5vw rgba(30, 58, 95, 0.05)",
          }}
        >
          <div
            style={{
              width: "6vw",
              height: "6vw",
              backgroundColor: "rgba(13, 148, 136, 0.08)",
              borderRadius: "50%",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              marginBottom: "3vh",
            }}
          >
            <div
              style={{
                width: "3.5vw",
                height: "3.5vw",
                backgroundColor: "#0D9488",
                borderRadius: "50%",
              }}
            />
          </div>

          <div
            style={{
              fontSize: "8vw",
              fontWeight: 800,
              color: "#0D9488",
              lineHeight: 1,
              letterSpacing: "-0.03em",
              marginBottom: "1vh",
            }}
          >
            ~70%
          </div>
          <div
            style={{
              fontSize: "1.4vw",
              fontWeight: 600,
              color: "#1E3A5F",
              marginBottom: "1.5vh",
            }}
          >
            complete on day one
          </div>
          <div
            style={{
              width: "4vw",
              height: "2px",
              backgroundColor: "#E2E8F0",
              marginBottom: "1.5vh",
            }}
          />
          <div
            style={{
              fontSize: "1.1vw",
              color: "#64748B",
              lineHeight: 1.5,
              maxWidth: "18vw",
            }}
          >
            Competitive integrators build all of this from zero; CP Inline starts at ~70% complete
          </div>
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
          <span>Page 4</span>
        </div>
      </div>
    </div>
  );
}
