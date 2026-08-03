export default function Slide3Features() {
  return (
    <div
      className="w-screen h-screen overflow-hidden relative"
      style={{
        backgroundColor: "#FAFBFC",
        fontFamily: "'Inter', sans-serif",
        padding: "4vh 4vw",
        boxSizing: "border-box",
        display: "grid",
        gridTemplateColumns: "1fr",
        gridTemplateRows: "auto 1fr auto",
        gap: "3vh",
        color: "#1E3A5F",
      }}
    >
      {/* Header */}
      <div
        style={{
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
          <div>CAPABILITIES</div>
          <div>2026</div>
        </div>
      </div>

      {/* Content */}
      <div style={{ display: "flex", flexDirection: "column", overflow: "hidden" }}>
        <div
          style={{
            fontSize: "1vw",
            fontWeight: 600,
            color: "#0D9488",
            marginBottom: "1vh",
            textTransform: "uppercase",
            letterSpacing: "0.06em",
          }}
        >
          Everything the line needs
        </div>
        <h2
          style={{
            fontSize: "3vw",
            fontWeight: 800,
            margin: "0 0 2.5vh 0",
            lineHeight: 1.1,
            letterSpacing: "-0.02em",
          }}
        >
          Everything the line needs
        </h2>

        {/* 2-column feature grid */}
        <div
          style={{
            display: "grid",
            gridTemplateColumns: "1fr 1fr",
            gap: "1.5vh 2.5vw",
            flex: 1,
          }}
        >
          {/* Left column — 4 items */}
          <div style={{ display: "flex", flexDirection: "column", gap: "1.5vh" }}>
            {/* Feature 1 */}
            <div
              style={{
                background: "#FFFFFF",
                borderRadius: "0.7vw",
                border: "1px solid #E2E8F0",
                borderLeft: "0.35vw solid #0D9488",
                padding: "1.5vh 1.5vw",
                boxShadow: "0 0.2vw 0.8vw rgba(30, 58, 95, 0.04)",
              }}
            >
              <div
                style={{
                  fontSize: "1.3vw",
                  fontWeight: 600,
                  color: "#1E3A5F",
                  marginBottom: "0.5vh",
                }}
              >
                Full-screen operator panel
              </div>
              <div
                style={{
                  fontSize: "1.1vw",
                  color: "#64748B",
                  lineHeight: 1.4,
                }}
              >
                Pass/fail banner readable at arm's length, live grade, GS1 AI breakdown
              </div>
            </div>

            {/* Feature 2 */}
            <div
              style={{
                background: "#FFFFFF",
                borderRadius: "0.7vw",
                border: "1px solid #E2E8F0",
                borderLeft: "0.35vw solid #0D9488",
                padding: "1.5vh 1.5vw",
                boxShadow: "0 0.2vw 0.8vw rgba(30, 58, 95, 0.04)",
              }}
            >
              <div
                style={{
                  fontSize: "1.3vw",
                  fontWeight: 600,
                  color: "#1E3A5F",
                  marginBottom: "0.5vh",
                }}
              >
                3–5 segment indicator pole
              </div>
              <div
                style={{
                  fontSize: "1.1vw",
                  color: "#64748B",
                  lineHeight: 1.4,
                }}
              >
                Grade-mapped steady/flash states — RED stop, AMBER warn, GREEN pass
              </div>
            </div>

            {/* Feature 3 */}
            <div
              style={{
                background: "#FFFFFF",
                borderRadius: "0.7vw",
                border: "1px solid #E2E8F0",
                borderLeft: "0.35vw solid #0D9488",
                padding: "1.5vh 1.5vw",
                boxShadow: "0 0.2vw 0.8vw rgba(30, 58, 95, 0.04)",
              }}
            >
              <div
                style={{
                  fontSize: "1.3vw",
                  fontWeight: 600,
                  color: "#1E3A5F",
                  marginBottom: "0.5vh",
                }}
              >
                Conveyor interrupt
              </div>
              <div
                style={{
                  fontSize: "1.1vw",
                  color: "#64748B",
                  lineHeight: 1.4,
                }}
              >
                Stops line on fail; default threshold 1 consecutive fail, lockable
              </div>
            </div>

            {/* Feature 4 */}
            <div
              style={{
                background: "#FFFFFF",
                borderRadius: "0.7vw",
                border: "1px solid #E2E8F0",
                borderLeft: "0.35vw solid #0D9488",
                padding: "1.5vh 1.5vw",
                boxShadow: "0 0.2vw 0.8vw rgba(30, 58, 95, 0.04)",
              }}
            >
              <div
                style={{
                  fontSize: "1.3vw",
                  fontWeight: 600,
                  color: "#1E3A5F",
                  marginBottom: "0.5vh",
                }}
              >
                GS1 DataMatrix UDI decode
              </div>
              <div
                style={{
                  fontSize: "1.1vw",
                  color: "#64748B",
                  lineHeight: 1.4,
                }}
              >
                GTIN, Expiry, Lot, Serial extracted every scan
              </div>
            </div>
          </div>

          {/* Right column — 3 items */}
          <div style={{ display: "flex", flexDirection: "column", gap: "1.5vh" }}>
            {/* Feature 5 */}
            <div
              style={{
                background: "#FFFFFF",
                borderRadius: "0.7vw",
                border: "1px solid #E2E8F0",
                borderLeft: "0.35vw solid #0D9488",
                padding: "1.5vh 1.5vw",
                boxShadow: "0 0.2vw 0.8vw rgba(30, 58, 95, 0.04)",
              }}
            >
              <div
                style={{
                  fontSize: "1.3vw",
                  fontWeight: 600,
                  color: "#1E3A5F",
                  marginBottom: "0.5vh",
                }}
              >
                Session and shift reports in Excel
              </div>
              <div
                style={{
                  fontSize: "1.1vw",
                  color: "#64748B",
                  lineHeight: 1.4,
                }}
              >
                Plus an unalterable flat-file event log
              </div>
            </div>

            {/* Feature 6 */}
            <div
              style={{
                background: "#FFFFFF",
                borderRadius: "0.7vw",
                border: "1px solid #E2E8F0",
                borderLeft: "0.35vw solid #0D9488",
                padding: "1.5vh 1.5vw",
                boxShadow: "0 0.2vw 0.8vw rgba(30, 58, 95, 0.04)",
              }}
            >
              <div
                style={{
                  fontSize: "1.3vw",
                  fontWeight: 600,
                  color: "#1E3A5F",
                  marginBottom: "0.5vh",
                }}
              >
                Full-frame JPEG archive on fail
              </div>
              <div
                style={{
                  fontSize: "1.1vw",
                  color: "#64748B",
                  lineHeight: 1.4,
                }}
              >
                Ready for offline DataMan Setup Tool (DMST) upload — onsite or remote regrading
              </div>
            </div>

            {/* Feature 7 */}
            <div
              style={{
                background: "#FFFFFF",
                borderRadius: "0.7vw",
                border: "1px solid #E2E8F0",
                borderLeft: "0.35vw solid #0D9488",
                padding: "1.5vh 1.5vw",
                boxShadow: "0 0.2vw 0.8vw rgba(30, 58, 95, 0.04)",
              }}
            >
              <div
                style={{
                  fontSize: "1.3vw",
                  fontWeight: 600,
                  color: "#1E3A5F",
                  marginBottom: "0.5vh",
                }}
              >
                RAID storage on your Win 11 PC
              </div>
              <div
                style={{
                  fontSize: "1.1vw",
                  color: "#64748B",
                  lineHeight: 1.4,
                }}
              >
                No cloud, no dependency, everything local
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Footer */}
      <div
        style={{
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
          <span>Page 3</span>
        </div>
      </div>
    </div>
  );
}
