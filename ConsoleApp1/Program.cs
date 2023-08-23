// See https://aka.ms/new-console-template for more information
Console.WriteLine("Hello, World!");
dynamic mySCAPI = SCAPI.Init();
//dynamic myModel = SCAPI.CreateNewModel(mySCAPI); // <-- Fails
dynamic myModel = SCAPI.OpenModel(mySCAPI,"C:\\Users\\User\\Desktop\\BigDMModels\\ITSKopie2020.erwin");
dynamic mySession = SCAPI.CreateSession(mySCAPI, myModel);
int i = 0;
if (mySession.IsOpen()) {
    dynamic myModelObjects = mySession.ModelObjects;
    Console.WriteLine("Model contains (" + myModelObjects.Count() + ")");
    foreach (dynamic myModelObject in myModelObjects) {
        Console.WriteLine(" - " + myModelObject.Name() + " " + myModelObject.ObjectId());


        if (i++ > 10)   // Quit out - too much to write
            break;
    }
}
Console.WriteLine("Done, World!");

public static class SCAPI {
    public static dynamic Init()
    {
        //const string progID = "SCAPI.Application";
        //Type foo = Type.GetTypeFromProgID (progID);

        var bar = Guid.Parse("6774E2C3-06E9-4943-A8D4-E3007AB1F42E");
        Type foo = Type.GetTypeFromCLSID(bar);

        dynamic COMobject = Activator.CreateInstance(foo);
        return COMobject;
    }

    public static dynamic CreateNewModel(dynamic scAppPtr)
    {
        dynamic scPUnitColPtr;
        scPUnitColPtr = scAppPtr.PersistenceUnits;
        object nil = Type.Missing;
        
        var bar = Guid.Parse("7D7B1602-9832-4AC6-A224-F0092FAF0D7E");
        Type foo = Type.GetTypeFromCLSID(bar);

        dynamic propBag = Activator.CreateInstance(foo);
        
        if (!propBag.Add("Name", "Test Model")) {
            Console.WriteLine("Fail 1");
            return null;
        }

        if (!propBag.Add("ModelType", 0)) {
            Console.WriteLine("Fail 2");
            return null;
        }
        dynamic scPUnitPtr = scPUnitColPtr.Create(propBag, nil); // <--- doesn't work - something about Name not being recognised as key??
        return scPUnitPtr;
    }
    public static dynamic OpenModel(dynamic scAppPtr, string fullModelPath)
    {
        dynamic scPUnitColPtr;
        dynamic scPUnitsColPtr = scAppPtr.PersistenceUnits;

        scPUnitColPtr = scPUnitsColPtr.Add(fullModelPath);
        return scPUnitColPtr;
    }
    public static dynamic CreateSession(dynamic scAppPtr, dynamic scPUnitColPtr)
    {
        dynamic scPSessionPtr;
        dynamic scPSessionsPtr = scAppPtr.Sessions;

        scPSessionPtr = scPSessionsPtr.Add();
        scPSessionPtr.Open(scPUnitColPtr, 0, 0);    // Level = SCD_SL_M0, Flags = SCD_SF_EXCLUSIVE = 1
        return scPSessionPtr;
    }
   
};

