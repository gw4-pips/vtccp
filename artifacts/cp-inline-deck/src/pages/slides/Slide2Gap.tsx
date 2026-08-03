export default function Slide2Gap() {
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
          <div>THE OPPORTUNITY</div>
          <div>2026</div>
        </div>
      </div>

      {/* Left — challenge side */}
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
          The gap we're closing
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
          The gap we're closing
        </h2>

        <div style={{ display: "flex", flexDirection: "column", gap: "2vh" }}>
          {/* Item 1 */}
          <div
            style={{
              display: "flex",
              gap: "1.5vw",
              alignItems: "flex-start",
              background: "#FFFFFF",
              padding: "2vh 2vw",
              borderRadius: "0.8vw",
              border: "1px solid #E2E8F0",
              boxShadow: "0 0.3vw 1vw rgba(30, 58, 95, 0.05)",
            }}
          >
            <div
              style={{
                fontSize: "1.1vw",
                fontWeight: 700,
                color: "#0D9488",
                backgroundColor: "rgba(13, 148, 136, 0.1)",
                width: "2.8vw",
                height: "2.8vw",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                borderRadius: "50%",
                flexShrink: 0,
              }}
            >
              1
            </div>
            <div>
              <div
                style={{
                  fontSize: "1.3vw",
                  fontWeight: 600,
                  color: "#1E3A5F",
                  marginBottom: "0.4vh",
                }}
              >
                You already own the verifier
              </div>
              <div
                style={{
                  fontSize: "1.1vw",
                  color: "#64748B",
                  lineHeight: 1.4,
                }}
              >
                A Cognex DM-475V-LBL high-speed inline unit
              </div>
            </div>
          </div>

          {/* Item 2 */}
          <div
            style={{
              display: "flex",
              gap: "1.5vw",
              alignItems: "flex-start",
              background: "#FFFFFF",
              padding: "2vh 2vw",
              borderRadius: "0.8vw",
              border: "1px solid #E2E8F0",
              boxShadow: "0 0.3vw 1vw rgba(30, 58, 95, 0.05)",
            }}
          >
            <div
              style={{
                fontSize: "1.1vw",
                fontWeight: 700,
                color: "#0D9488",
                backgroundColor: "rgba(13, 148, 136, 0.1)",
                width: "2.8vw",
                height: "2.8vw",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                borderRadius: "50%",
                flexShrink: 0,
              }}
            >
              2
            </div>
            <div>
              <div
                style={{
                  fontSize: "1.3vw",
                  fontWeight: 600,
                  color: "#1E3A5F",
                  marginBottom: "0.4vh",
                }}
              >
                Full industrial integration is beyond current budget and timeline
              </div>
              <div
                style={{
                  fontSize: "1.1vw",
                  color: "#64748B",
                  lineHeight: 1.4,
                }}
              >
                A desktop stand adapted for conveyor-side use gets the line running now
              </div>
            </div>
          </div>

          {/* Item 3 */}
          <div
            style={{
              display: "flex",
              gap: "1.5vw",
              alignItems: "flex-start",
              background: "#FFFFFF",
              padding: "2vh 2vw",
              borderRadius: "0.8vw",
              border: "1px solid #E2E8F0",
              boxShadow: "0 0.3vw 1vw rgba(30, 58, 95, 0.05)",
            }}
          >
            <div
              style={{
                fontSize: "1.1vw",
                fontWeight: 700,
                color: "#0D9488",
                backgroundColor: "rgba(13, 148, 136, 0.1)",
                width: "2.8vw",
                height: "2.8vw",
                display: "flex",
                alignItems: "center",
                justifyContent: "center",
                borderRadius: "50%",
                flexShrink: 0,
              }}
            >
              3
            </div>
            <div>
              <div
                style={{
                  fontSize: "1.3vw",
                  fontWeight: 600,
                  color: "#1E3A5F",
                  marginBottom: "0.4vh",
                }}
              >
                CP Inline delivers production-grade control
              </div>
              <div
                style={{
                  fontSize: "1.1vw",
                  color: "#64748B",
                  lineHeight: 1.4,
                }}
              >
                At a fraction of a full Cognex integration cost
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Right — solution card */}
      <div
        style={{
          display: "flex",
          flexDirection: "column",
          justifyContent: "center",
        }}
      >
        <div
          style={{
            background: "#FFFFFF",
            padding: "4vh 3vw",
            borderRadius: "1vw",
            border: "1px solid #E2E8F0",
            height: "100%",
            display: "flex",
            flexDirection: "column",
            justifyContent: "center",
            boxSizing: "border-box",
            boxShadow: "0 0.5vw 1.5vw rgba(30, 58, 95, 0.05)",
          }}
        >
          <div
            style={{
              fontSize: "0.9vw",
              fontWeight: 600,
              color: "#0D9488",
              textTransform: "uppercase",
              letterSpacing: "0.06em",
              marginBottom: "2vh",
            }}
          >
            Out of the box
          </div>
          <div
            style={{
              fontSize: "1.8vw",
              fontWeight: 700,
              color: "#1E3A5F",
              lineHeight: 1.2,
              marginBottom: "3vh",
              borderBottom: "1px solid #E2E8F0",
              paddingBottom: "3vh",
            }}
          >
            UDI GS1 DataMatrix decode, grade-based pass/fail, and archiving
          </div>

          <div style={{ display: "flex", flexDirection: "column", gap: "2.5vh" }}>
            <div style={{ display: "flex", gap: "1vw", alignItems: "flex-start" }}>
              <div
                style={{
                  width: "0.8vw",
                  height: "0.8vw",
                  backgroundColor: "#0D9488",
                  borderRadius: "50%",
                  marginTop: "0.4vh",
                  flexShrink: 0,
                }}
              />
              <div style={{ fontSize: "1.2vw", color: "#475569", lineHeight: 1.5 }}>
                Grade-based pass/fail on every scan — automatically
              </div>
            </div>
            <div style={{ display: "flex", gap: "1vw", alignItems: "flex-start" }}>
              <div
                style={{
                  width: "0.8vw",
                  height: "0.8vw",
                  backgroundColor: "#0D9488",
                  borderRadius: "50%",
                  marginTop: "0.4vh",
                  flexShrink: 0,
                }}
              />
              <div style={{ fontSize: "1.2vw", color: "#475569", lineHeight: 1.5 }}>
                GTIN, Expiry, Lot, Serial extracted from every GS1 DataMatrix
              </div>
            </div>
            <div style={{ display: "flex", gap: "1vw", alignItems: "flex-start" }}>
              <div
                style={{
                  width: "0.8vw",
                  height: "0.8vw",
                  backgroundColor: "#0D9488",
                  borderRadius: "50%",
                  marginTop: "0.4vh",
                  flexShrink: 0,
                }}
              />
              <div style={{ fontSize: "1.2vw", color: "#475569", lineHeight: 1.5 }}>
                Image archive of every fail — ready for offline DMST review
              </div>
            </div>
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
          <span>Page 2</span>
        </div>
      </div>
    </div>
  );
}
