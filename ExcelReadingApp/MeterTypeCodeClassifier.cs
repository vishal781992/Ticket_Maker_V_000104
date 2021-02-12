using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReadingApp
{
    class MeterTypeCodeClassifier
    {
        public string MeterTypeCode_AppenderString = string.Empty;
        public string MeterTypeCode001(string MeterTypeCode)
        {
            string Position1 = MeterTypeCode.Substring(0, 1);
            MeterTypeCode_AppenderString = string.Empty; //clean this before you start
            try
            {
                int.TryParse(Position1, out int Position1_int);
                switch(Position1_int)
                {
                    case 1:
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "Form 1S" + ",";
                        break;

                    case 2:
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "Form 2S" + ",";
                        break;

                    case 3:
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "Form 3S" + ",";
                        break;

                    case 4:
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "Form 4S" + ",";
                        break;

                    case 5:
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "Form 5S" + ",";
                        break;

                    case 6:
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "Form 6S" + ",";
                        break;

                    case 8:
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "Form 8S" + ",";
                        break;

                    case 9:
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "Form 9S" + ",";
                        break;

                    default:
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "--" + ",";
                        break;
                }
            }
            catch { MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "-c-" + ","; }
            MeterTypeCode002(MeterTypeCode);

            return MeterTypeCode_AppenderString;
        }

        public void MeterTypeCode002(string MeterTypeCode)
        {
            string Position2 = MeterTypeCode.Substring(1, 1);
            try
            {
                switch (Position2.ToUpper())
                {
                    case "A":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "120V/100A" + ",";
                        break;

                    case "B":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "120V/200A" + ",";
                        break;

                    case "C":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "120V/320A" + ",";
                        break;

                    case "D":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "120V/20A" + ",";
                        break;

                    case "E":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "240V/200A" + ",";
                        break;

                    case "F":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "240V/320A" + ",";
                        break;

                    case "G":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "240V/20A" + ",";
                        break;

                    case "H":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "480V/200A" + ",";
                        break;

                    case "J":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "480V/320A" + ",";
                        break;

                    case "K":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "480V/20A" + ",";
                        break;

                    case "L":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "120-480V/200A" + ",";
                        break;

                    case "M":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "120-480V/320A" + ",";
                        break;

                    case "N":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "120-480V/20A" + ",";
                        break;

                    case "P":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "120-480V/100A" + ",";
                        break;

                    case "Q":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "480V/100A" + ",";
                        break;

                    case "R":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "240V/100A" + ",";
                        break;

                }
            }
            catch { }
            MeterTypeCode004(MeterTypeCode);
        }

        public void  MeterTypeCode004(string MeterTypeCode)
        {
            string Position2 = MeterTypeCode.Substring(3, 1);
            try
            {
                switch (Position2.ToUpper())
                {
                    case "A":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "No Comm." + ".";
                        break;

                    case "B":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "HPRadio" + ".";
                        break;

                    case "C":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "PulseOP.formC" + ".";
                        break;

                    case "D":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "pulseOP.formA" + ".";
                        break;

                    case "E":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "RS485" + ".";
                        break;

                    case "F":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "RF/Pulse.C" + ".";
                        break;

                    case "G":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "RF/Pulse.A" + ".";
                        break;

                    case "H":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "LoraDualRCV" + ".";
                        break;

                    case "I":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "GridWide" + ".";
                        break;

                    case "J":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "DoD.Radio" + ".";
                        break;

                    case "K":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "visionLTE" + ".";
                        break;

                    case "L":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "LocusEnergy" + ".";
                        break;

                    case "M":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "LocusLTE" + ".";
                        break;

                    case "N":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "NexGrid" + ".";
                        break;

                    case "O":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "CatM1.Sprint" + ".";
                        break;

                    case "P":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "CatM1.Sierra" + ".";
                        break;

                    case "Q":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "ModBus" + ".";
                        break;

                    case "R":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "Optical" + ".";
                        break;

                    case "S":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "HPAirpoint2" + ".";
                        break;

                    case "T":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "HPAirpoint3" + ".";
                        break;

                    case "U":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "CatM1.LTE" + ".";
                        break;

                    case "V":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "HPAirpoint4" + ".";
                        break;

                    case "W":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "HPAirpoint5" + ".";
                        break;

                    case "X":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "VisionLTE.R200" + ".";
                        break;

                    case "Y":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "VisionLTE.L201" + ".";
                        break;

                    case "Z":
                        MeterTypeCode_AppenderString = MeterTypeCode_AppenderString + "LoRa" + ".";
                        break;

                }
            }
            catch { }

        }
    }
}
