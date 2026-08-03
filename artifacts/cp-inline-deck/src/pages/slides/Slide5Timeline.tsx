export default function Slide5Timeline() {
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
          <div>DEPLOYMENT TIMELINE</div>
          <div>2026</div>
        </div>
      </div>

      {/* Content */}
      <div
        style={{
          display: "flex",
          flexDirection: "column",
          gap: "2.5vh",
          overflow: "hidden",
        }}
      >
        {/* Headline */}
        <div>
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
            6 weeks to go-live
          </div>
          <h2
            style={{
              fontSize: "3vw",
              fontWeight: 800,
              margin: 0,
              lineHeight: 1.1,
              letterSpacing: "-0.02em",
            }}
          >
            6 weeks to go-live
          </h2>
        </div>

        {/* Timeline cards — 4 phases */}
        <div
          style={{
            display: "flex",
            gap: "1.5vw",
            flex: 1,
          }}
        >
          {/* Phase 1 — Weeks 1–3 */}
          <div
            style={{
              flex: "1.5",
              background: "#FFFFFF",
              borderRadius: "0.8vw",
              border: "1px solid #E2E8F0",
              padding: "2.5vh 2vw",
              display: "flex",
              flexDirection: "column",
              boxShadow: "0 0.3vw 1vw rgba(30, 58, 95, 0.05)",
              boxSizing: "border-box",
            }}
          >
            <div
              style={{
                fontSize: "0.9vw",
                fontWeight: 600,
                color: "#0D9488",
                textTransform: "uppercase",
                letterSpacing: "0.06em",
                marginBottom: "1vh",
              }}
            >
              Weeks 1–3
            </div>
            <div
              style={{
                fontSize: "1.6vw",
                fontWeight: 700,
                color: "#1E3A5F",
                marginBottom: "1.5vh",
              }}
            >
              Build and Test
            </div>
            <div
              style={{
                width: "3vw",
                height: "3px",
                backgroundColor: "rgba(13, 148, 136, 0.3)",
                borderRadius: "2px",
                marginBottom: "1.5vh",
              }}
            />
            <div
              style={{
                fontSize: "1.1vw",
                color: "#64748B",
                lineHeight: 1.5,
              }}
            >
              Relay I/O, full operator panel, reporting, installer
            </div>
          </div>

          {/* Phase 2 — Week 4 */}
          <div
            style={{
              flex: 1,
              background: "#FFFFFF",
              borderRadius: "0.8vw",
              border: "1px solid #E2E8F0",
              padding: "2.5vh 2vw",
              display: "flex",
              flexDirection: "column",
              boxShadow: "0 0.3vw 1vw rgba(30, 58, 95, 0.05)",
              boxSizing: "border-box",
            }}
          >
            <div
              style={{
                fontSize: "0.9vw",
                fontWeight: 600,
                color: "#0D9488",
                textTransform: "uppercase",
                letterSpacing: "0.06em",
                marginBottom: "1vh",
              }}
            >
              Week 4
            </div>
            <div
              style={{
                fontSize: "1.6vw",
                fontWeight: 700,
                color: "#1E3A5F",
                marginBottom: "1.5vh",
              }}
            >
              PE Trial
            </div>
            <div
              style={{
                width: "3vw",
                height: "3px",
                backgroundColor: "rgba(13, 148, 136, 0.3)",
                borderRadius: "2px",
                marginBottom: "1.5vh",
              }}
            />
            <div
              style={{
                fontSize: "1.1vw",
                color: "#64748B",
                lineHeight: 1.5,
              }}
            >
              First live product on the customer line
            </div>
          </div>

          {/* Phase 3 — Week 5 */}
          <div
            style={{
              flex: 1,
              background: "#FFFFFF",
              borderRadius: "0.8vw",
              border: "1px solid #E2E8F0",
              padding: "2.5vh 2vw",
              display: "flex",
              flexDirection: "column",
              boxShadow: "0 0.3vw 1vw rgba(30, 58, 95, 0.05)",
              boxSizing: "border-box",
            }}
          >
            <div
              style={{
                fontSize: "0.9vw",
                fontWeight: 600,
                color: "#0D9488",
                textTransform: "uppercase",
                letterSpacing: "0.06em",
                marginBottom: "1vh",
              }}
            >
              Week 5
            </div>
            <div
              style={{
                fontSize: "1.6vw",
                fontWeight: 700,
                color: "#1E3A5F",
                marginBottom: "1.5vh",
              }}
            >
              Hardening
            </div>
            <div
              style={{
                width: "3vw",
                height: "3px",
                backgroundColor: "rgba(13, 148, 136, 0.3)",
                borderRadius: "2px",
                marginBottom: "1.5vh",
              }}
            />
            <div
              style={{
                fontSize: "1.1vw",
                color: "#64748B",
                lineHeight: 1.5,
              }}
            >
              Operator training, report sign-off, grade threshold tuning
            </div>
          </div>

          {/* Phase 4 — Week 6 */}
          <div
            style={{
              flex: 1,
              background: "#0D9488",
              borderRadius: "0.8vw",
              padding: "2.5vh 2vw",
              display: "flex",
              flexDirection: "column",
              boxShadow: "0 0.5vw 1.5vw rgba(13, 148, 136, 0.25)",
              boxSizing: "border-box",
            }}
          >
            <div
              style={{
                fontSize: "0.9vw",
                fontWeight: 600,
                color: "rgba(255,255,255,0.7)",
                textTransform: "uppercase",
                letterSpacing: "0.06em",
                marginBottom: "1vh",
              }}
            >
              Week 6
            </div>
            <div
              style={{
                fontSize: "1.6vw",
                fontWeight: 700,
                color: "#FFFFFF",
                marginBottom: "1.5vh",
              }}
            >
              Go-Live
            </div>
            <div
              style={{
                width: "3vw",
                height: "3px",
                backgroundColor: "rgba(255,255,255,0.3)",
                borderRadius: "2px",
                marginBottom: "1.5vh",
              }}
            />
            <div
              style={{
                fontSize: "1.1vw",
                color: "rgba(255,255,255,0.85)",
                lineHeight: 1.5,
              }}
            >
              Live production go-live
            </div>
          </div>
        </div>

        {/* Prerequisites panel */}
        <div
          style={{
            background: "#FFFFFF",
            borderRadius: "0.8vw",
            border: "1px solid #E2E8F0",
            padding: "2vh 2.5vw",
            display: "flex",
            alignItems: "center",
            gap: "3vw",
            boxShadow: "0 0.2vw 0.8vw rgba(30, 58, 95, 0.04)",
          }}
        >
          <div style={{ flexShrink: 0 }}>
            <div
              style={{
                fontSize: "0.9vw",
                fontWeight: 600,
                color: "#0D9488",
                textTransform: "uppercase",
                letterSpacing: "0.06em",
                marginBottom: "0.5vh",
              }}
            >
              Before Week 1
            </div>
            <div
              style={{
                fontSize: "1.2vw",
                fontWeight: 700,
                color: "#1E3A5F",
              }}
            >
              Two engineering decisions needed
            </div>
          </div>
          <div
            style={{
              width: "1px",
              alignSelf: "stretch",
              backgroundColor: "#E2E8F0",
            }}
          />
          <div style={{ display: "flex", gap: "3vw", flex: 1 }}>
            <div style={{ display: "flex", gap: "0.8vw", alignItems: "center" }}>
              <div
                style={{
                  width: "0.7vw",
                  height: "0.7vw",
                  backgroundColor: "#0D9488",
                  borderRadius: "50%",
                  flexShrink: 0,
                }}
              />
              <div style={{ fontSize: "1.15vw", color: "#475569", lineHeight: 1.4 }}>
                Relay board model and conveyor interrupt circuit spec
              </div>
            </div>
            <div style={{ display: "flex", gap: "0.8vw", alignItems: "center" }}>
              <div
                style={{
                  width: "0.7vw",
                  height: "0.7vw",
                  backgroundColor: "#0D9488",
                  borderRadius: "50%",
                  flexShrink: 0,
                }}
              />
              <div style={{ fontSize: "1.15vw", color: "#475569", lineHeight: 1.4 }}>
                IMAGE.SEND firmware confirmation on the DM-475V unit
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
          <span>Page 5</span>
        </div>
      </div>
    </div>
  );
}
