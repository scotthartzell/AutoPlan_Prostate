using System;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Reflection;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

// TODO: Replace the following version attributes by creating AssemblyInfo.cs. You can do this in the properties of the Visual Studio project.
[assembly: AssemblyVersion("1.0.1.2")]
[assembly: AssemblyFileVersion("1.0.0.1")]
[assembly: AssemblyInformationalVersion("1.0")]

// TODO: Uncomment the following line if the script requires write access.
[assembly: ESAPIScript(IsWriteable = true)]

namespace AutoPlan_Prostate
{
    class Program
    {
        private static string _patientId;
        private static string _imageId;
        private static string _strSetId;

        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                if(args.Any())
                {
                    _patientId = args.First().Split(';').First();
                    _imageId = args.First().Split(';').ElementAt(1);
                    _strSetId = args.First().Split(';').Last();
                }
                using (Application app = Application.CreateApplication())
                {
                    Execute(app);
                }
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e.ToString());
            }
            Console.ReadLine(); // this will let the console remain open if crash output returned
        }
        static void Execute(Application app)
        {
            // TODO: Add your code here.
            // open the patient, use the launcher to pass into the application, have user type in
            //if(String.IsNullOrEmpty(_patientId))
            //{
            //    Console.WriteLine("Please enter patient ID");
            //    _patientId = Console.ReadLine();
            //}
            if (String.IsNullOrEmpty(_patientId))
            {
                AskUserId("Patient Id", out _patientId);
            }
            Patient patient = app.OpenPatientById(_patientId);
            if (patient == null) { Console.WriteLine("Incorrect patient ID"); return; }
            Console.WriteLine($"The open patient is: {patient.Name}");
            patient.BeginModifications(); // this is a must if making any patient modifications

            // create course, but not every time, just for the automated plans
            string courseId = "AutoCourse";
            Course course = null;
            if(patient.Courses.Any(c => c.Id.Equals(courseId)))
            {
                course = patient.Courses.First(c => c.Id.Equals(courseId));
            }
            else
            {
                course = patient.AddCourse();
                course.Id = courseId;
            }
            // now create the plan, polymorphism, PlanSetup commonalities, external plan setup class is needed for creation
            //string imageId = "CT_2";
            //string strSetId = "CT_1";
            if (String.IsNullOrEmpty(_imageId))
            {
                AskUserId("Image Id", out _imageId);
            }
            //else { _imageId = "CT_2"; }
            if (String.IsNullOrEmpty(_strSetId))
            {
                AskUserId("Structure Set Id", out _strSetId);
            }
            //else { _strSetId = "CT_1"; }

            StructureSet structureSet = patient.StructureSets.First(ss => ss.Id.Equals(_strSetId) && ss.Image.Id.Equals(_imageId));
            ExternalPlanSetup plan = course.AddExternalPlanSetup(structureSet);
            Console.WriteLine($"Plan {plan.Id} create on course {course.Id}");

            //ExternalPlanSetup plan = course.AddExternalPlanSetup(structureSet)

            double[] gantryAngles = new double[] {220,255,290,325,25,55,95,130};
            ExternalBeamMachineParameters parameters = new ExternalBeamMachineParameters("HESN5","10X",600,"STATIC",null);
            // new new VVector());
            foreach (var gantryAngle in gantryAngles)
            {
                plan.AddStaticBeam(parameters,
                    new VRect<double>(-50, -50, 50, 50),
                    0,
                    gantryAngle,
                    0,
                    structureSet.Image.UserOrigin);
            }
            Console.WriteLine($"Created {plan.Beams.Count()} beams");
            //set the target and the Rx
            string targetId = "PTVprost SV marg";
            Structure target = structureSet.Structures.First(st => st.Id.Equals(targetId));
            StringBuilder errorString = new StringBuilder();
            //if (!plan.SetTargetStructureIfNoDose(target, errorString))
            //{
            //    Console.WriteLine($"Could not set target:\n{errorString}");
            //};
            plan.SetTargetStructureIfNoDose(target, errorString);
            plan.SetPrescription(28, new DoseValue(250, DoseValue.DoseUnit.cGy), 1.0);
            Console.WriteLine($"Plan prescribed {plan.TotalDose} in {plan.NumberOfFractions} fractions to {target.Id}");
            // rapid plan time, in 16.1 don't need names of models
            int rpCount = 0;
            foreach(var rp in app.Calculation.GetDvhEstimationModelSummaries())// new to 16.1
            {
                Console.WriteLine($"[{rpCount}].\t{rp.Name} - {rp.TreatmentSite}");
                rpCount++;
            }
            Console.WriteLine("Please select a RapidPlan model");
            int rpSelect = Convert.ToInt32(Console.ReadLine());
            var rpModel = app.Calculation.GetDvhEstimationModelSummaries().ElementAt(rpSelect);

            Dictionary<string, string> structureMatches = new Dictionary<string, string>();
            Dictionary<string, DoseValue> targetMatches = new Dictionary<string, DoseValue>();
            foreach(var rpStructure in app.Calculation.GetDvhEstimationModelStructures(rpModel.ModelUID))
            {
                if(rpStructure.StructureType == DVHEstimationStructureType.PTV)
                {
                    targetMatches.Add(target.Id, plan.TotalDose);
                    structureMatches.Add(target.Id, rpStructure.Id);
                }
                else
                {
                    if(structureSet.Structures.Any(st => st.Id.Equals(rpStructure.Id)))
                    {
                        Structure s = structureSet.Structures.First(st => st.Id.Equals(rpStructure.Id));
                        structureMatches.Add(s.Id, rpStructure.Id);
                    }
                }
            }
            Console.WriteLine("Setting DVH Estimates...");
            plan.CalculateDVHEstimates(rpModel.Name, targetMatches, structureMatches);
            //add NTO
            plan.OptimizationSetup.AddAutomaticNormalTissueObjective(100);
            //create ring
            string ringId = "NS_Ring05";
            Structure ring = null;
            if (structureSet.Structures.Any(st => st.Id.Equals(ringId, StringComparison.OrdinalIgnoreCase)))
            {
                ring = structureSet.Structures.First(st => st.Id.Equals(ringId, StringComparison.OrdinalIgnoreCase));
            }
            else
            {
                ring = structureSet.AddStructure("CONTROL", ringId);
                ring.SegmentVolume = target.SegmentVolume.Margin(5).Sub(target.SegmentVolume);
            }
            plan.OptimizationSetup.AddPointObjective(ring, OptimizationObjectiveOperator.Upper, plan.TotalDose * 1.02, 0, 100);
            Console.WriteLine("Optimizing...");
            plan.Optimize();
            Console.WriteLine("Calculating Leaf Motions...");
            plan.CalculateLeafMotions();

            Console.WriteLine("Calculating Dose...");
            plan.CalculateDose();
            Console.WriteLine("Saving Plan...");
            app.SaveModifications();
        }
        // this will look for the active (drop to view) ct and structure set, need this to be static for finding open patient
        private static void AskUserId(string desc, out string id)
        {
            Console.WriteLine($"Please enter the {desc}");
            id = Console.ReadLine();
        }
    }
}
