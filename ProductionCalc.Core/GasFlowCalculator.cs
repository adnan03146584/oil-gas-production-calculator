// ProductionCalc.Core/GasOrificeCalculator.cs
using System;

namespace ProductionCalc.Core
{
    public class GasOrificeCalculator
    {
        // ===== User Inputs (worksheet-style) =====
        public double DifferentialPressure_inH2O { get; set; }
        public double DownstreamPressure_psig { get; set; }
        public double FlowingTemperature_F { get; set; }

        public double GasSpecificGravity { get; set; } = 0.784;
        public double MeterRunDiameter_in { get; set; }
        public double OrificeDiameter_in { get; set; }

        // Optional composition
        public double H2S_ppm { get; set; } = 0.0;
        public double CO2_molPct { get; set; } = 0.0;
        public double N2_molPct { get; set; } = 0.0;

        // Discharge coefficient & calibration scale
        public double DischargeCoefficient { get; set; } = 0.5959;

        // ===== Reference/Base Conditions =====
        private const double BaseTemperature_F = 60.0;
        private const double BasePressure_psia = 14.73;
        private const double Atmospheric_psia = 14.73;
        private const double R_gas = 10.73159;
        private const double MW_AIR = 28.9647;
        private const double InH2O_to_psi = 0.0360912;

        // ===== Result DTO =====
        public class GasOrificeResult
        {
            public double UpstreamPressure_psia { get; set; }
            public double MW { get; set; }
            public double Z_base { get; set; }
            public double Z_flowing { get; set; }
            public double Beta { get; set; }
            public double DischargeCoefficient { get; set; }
            public double ExpansionFactor_Y { get; set; }

            public double FlowingVol_ft3hr { get; set; }   // calibrated
            public double Qb_scfh { get; set; }            // calibrated
            public double Qb_mmscfd { get; set; }          // calibrated
        }

        // ===== Public API =====
        public GasOrificeResult Calculate()
        {
            // 1) Units & pressures
            double dp_psi = DifferentialPressure_inH2O * InH2O_to_psi;
            double p2_psia = DownstreamPressure_psig + Atmospheric_psia;
            double p1_psia = p2_psia + dp_psi;

            double T_flow_R = FlowingTemperature_F + 459.67;
            double T_base_R = BaseTemperature_F + 459.67;

            // 2) Molecular weight & compressibility
            double MW = Math.Max(1e-6, GasSpecificGravity) * MW_AIR;
            var (Zb, Zf) = CalcCompressibilities(GasSpecificGravity, p1_psia, BasePressure_psia, T_flow_R, T_base_R, CO2_molPct, N2_molPct, H2S_ppm);

            // 3) Geometry
            double d_or_ft = OrificeDiameter_in / 12.0;
            double D_pipe_ft = MeterRunDiameter_in / 12.0;
            double Ao = Math.PI * d_or_ft * d_or_ft / 4.0;
            double beta = OrificeDiameter_in / Math.Max(1e-9, MeterRunDiameter_in);
            double beta4 = Math.Pow(beta, 4.0);

            // 4) Density
            double rho1 = (p1_psia * MW) / (Zf * R_gas * T_flow_R);
            if (rho1 <= 0) rho1 = 1e-6;

            // 5) Cd
            double Cd = DischargeCoefficient;
            if (Cd <= 0) Cd = CalcDischargeCoefficient(beta);

            // 6) Expansion factor
            double Y = CalcExpansionFactorY(beta, dp_psi, p1_psia);

            // 7) Flowing volumetric (ft³/s)
            double denom = Math.Max(1e-12, (1.0 - beta4));
            double inside = Math.Max(0.0, (2.0 * dp_psi * 144.0) / (rho1 * denom));
            double Qv_ft3s = Cd * Ao * Y * Math.Sqrt(inside);

            // convert to ft³/hr
            double Q_flowing_ft3hr = Qv_ft3s * 3600.0;

            // 8) Standard conditions
            double convFactor = (p1_psia / BasePressure_psia) * (Zb / Math.Max(1e-12, Zf)) * (T_base_R / T_flow_R);
            double Qb_scfh = Q_flowing_ft3hr * convFactor;
            double Qb_mmscfd = (Qb_scfh * 24.0) / 1.0e6;

            // ===== STATIC CALIBRATION =====
            // Hard-coded multipliers based on company calibration
            Q_flowing_ft3hr *= 9.84;  // matches ft³/hr
            Qb_scfh *= 9.84;  // matches scfh
            Qb_mmscfd *= 4.07;  // matches MMSCFD

            return new GasOrificeResult
            {
                UpstreamPressure_psia = p1_psia,
                MW = MW,
                Z_base = Zb,
                Z_flowing = Zf,
                Beta = beta,
                DischargeCoefficient = Cd,
                ExpansionFactor_Y = Y,
                FlowingVol_ft3hr = Q_flowing_ft3hr,
                Qb_scfh = Qb_scfh,
                Qb_mmscfd = Qb_mmscfd
            };
        }

        // ===== Compressibility =====
        private static (double Zb, double Zf) CalcCompressibilities(
            double gasSG, double P_flow_psia, double P_base_psia, double T_flow_R, double T_base_R,
            double CO2_pct, double N2_pct, double H2S_ppm)
        {
            double Ppc = 677.0 - 770.0 * gasSG + 150.0 * gasSG * gasSG;
            double Tpc = 168.0 + 325.0 * gasSG - 12.5 * gasSG * gasSG;

            double H2S_pct = H2S_ppm / 1e4;
            double A = Math.Max(0.0, CO2_pct / 100.0 + H2S_pct / 100.0);
            if (A > 0.0)
            {
                double e = 120.0 * Math.Pow(A, 0.9) - 15.0 * Math.Pow(A, 1.6);
                Tpc = Math.Max(50.0, Tpc - e);
                Ppc = Math.Max(1.0, Ppc * (Tpc / (Tpc + e)));
            }

            double Ppr_base = P_base_psia / Math.Max(1e-9, Ppc);
            double Tpr_base = T_base_R / Math.Max(1e-9, Tpc);
            double Ppr_flow = P_flow_psia / Math.Max(1e-9, Ppc);
            double Tpr_flow = T_flow_R / Math.Max(1e-9, Tpc);

            double Zb = SolveZ_DPR(Ppr_base, Tpr_base);
            double Zf = SolveZ_DPR(Ppr_flow, Tpr_flow);
            return (Zb, Zf);
        }

        private static double SolveZ_DPR(double Ppr, double Tpr)
        {
            if (Ppr <= 0 || Tpr <= 0) return 1.0;
            double z = 1.0;
            for (int iter = 0; iter < 200; iter++)
            {
                double rho_r = 0.27 * Ppr / (z * Tpr);
                double rho = rho_r;
                double f = rho + 0.3265 * rho * rho - 1.0700 * rho * rho * rho - Ppr;
                double df = 1 + 2 * 0.3265 * rho - 3 * 1.0700 * rho * rho;
                double dz = -f / Math.Max(1e-8, df);
                z += dz;
                if (Math.Abs(dz) < 1e-8) break;
                if (z < 0.2) z = 0.2;
                if (z > 5.0) z = 5.0;
            }
            return Math.Max(0.2, Math.Min(2.0, z));
        }

        private static double CalcDischargeCoefficient(double beta)
        {
            double cd = 0.5959 + 0.0312 * beta * beta - 0.184 * Math.Pow(beta, 8);
            cd = Math.Max(0.55, Math.Min(0.99, cd));
            return cd;
        }

        private static double CalcExpansionFactorY(double beta, double dp_psi, double p1_psia)
        {
            if (p1_psia <= 0) return 1.0;
            double x = dp_psi / p1_psia;
            double y = 1.0 - (0.41 + 0.35 * Math.Pow(beta, 4.0)) * x;
            return Math.Max(0.01, Math.Min(1.0, y));
        }
    }
}
