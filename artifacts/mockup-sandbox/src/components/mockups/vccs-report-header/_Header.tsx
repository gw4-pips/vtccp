import "./_group.css";

type HeaderProps = {
  proposed?: boolean;
};

export function Header({ proposed = false }: HeaderProps) {
  return (
    <div className="vccs-preview-stage">
      <div className={`vccs-report-page ${proposed ? "vccs-proposed" : ""}`}>
        <header className="vccs-header">
          <div className="vccs-header-row">
            <div className="vccs-logo-box">
              <img
                className="vccs-logo"
                src="/__mockup/images/vccs-header/vccs_logo.png"
                alt="VCCS"
              />
            </div>
            <div className="vccs-header-meta">
              <div className="device">Device: Webscan TruCheck</div>
              <div className="line">Serial: TC-823-0610-028</div>
              <div className="line">Software: 3.03.74</div>
            </div>
            <div className="vccs-header-title">
              <h1>
                VCCS RFID <em>VeriWedge™ PowerPro</em>
                <br />
                Barcode-to-RFID Validation Report
              </h1>
              <div className="vccs-rfid-badge">✓ RFID MATCHED</div>
            </div>
            <div className="vccs-company-box">
              <img
                className="vccs-company-logo"
                src="/__mockup/images/vccs-header/pips_logo.png"
                alt="PIPS"
              />
            </div>
          </div>
          <div className="vccs-header-copyright">
            www.vccs.llc — © Copyright 2026 All rights reserved.{" "}
            <em>
              VCCS RFID VeriWedge™ is a product of Verifier Calibration &amp;
              Conformance Solutions, LLC
            </em>
          </div>
        </header>
      </div>
    </div>
  );
}
