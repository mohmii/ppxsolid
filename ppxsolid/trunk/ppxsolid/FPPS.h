#define DllExport __declspec( dllexport )

bool DllExport InitUserDLL3() ;
void DllExport Draft_Separation();
void DllExport File_Open();
//void DllExport TRV_Extraction();
bool DllExport TRV_Extraction();
long DllExport Surface_Generation_parallel_to_XY();
long DllExport Surface_Generation_forSlope();
void DllExport TRV_Partition();	
void DllExport Body_Classification();	
void DllExport Sub_Body_Reconfiguration();
void DllExport Add_OpenFace();
void DllExport Order_Decision();
void DllExport Order_Decision2();
void DllExport Feature_Recognition();
void DllExport Feature_Recognition2();

void DllExport RolledSteel_and_CarbonSteel();
void DllExport AlloyedSteel_and_ToolSteel();
void DllExport Stainless_Steel();


void DllExport Tool_CuttingPattern_Decision();
void DllExport Estimate_MachiningCost();
void DllExport Estimate_MachiningCost2();
void DllExport Assess_PluralProposals();
void DllExport Assess_PluralProposals2(BSTR);

void DllExport CreateAttribute0();
void DllExport CreateAttribute1();
void DllExport CreateAttribute2();
void DllExport CreateAttribute3();

void DllExport Machining_by_NormalMC();
void DllExport Machining_by_HighSpeedMC();

void DllExport Main_Process();
void DllExport Get_PartionPattern();
void DllExport Sub_ProcessPlanning_and_OperationPlanning();

void DllExport Viewpoint_for_Figure();
void DllExport Reset_System();
void DllExport Save_Candidate();
void DllExport Load_Candidate();
void DllExport ChangeCurrentDirectory(BSTR);
void DllExport Set_MTMCT(BSTR);

void DllExport Assess_PluralProposals3();

//--------------------------------------------ééçÏ
void DllExport TRV_Extraction2();
void DllExport Body_Classification2();
long DllExport Surface_Generation_parallel_to_XY2();
long DllExport Surface_Generation_forSlope2();
long DllExport Sub_Body_Reconfiguration2();
void DllExport Make_IntermediateShape_of_Workpeace(BSTR);
bool DllExport Remove_Fillet_TRV();
long DllExport Surface_Generation_byEdge();
void DllExport TRV_Partition_bySGE();
void DllExport Surface_Identity();
//--------------------------------------------

/*
BOOL DllExport  InitUserDLL() ;
void DllExport StartDlg();
*/
//çﬁóøéÌóﬁä«óùópÇÃíËêîåQ
#define MT_RS_CS	1
#define MT_AS_TS	2
#define MT_SS		3

//MCéÌóﬁä«óùópÇÃíËêîåQ
#define MCT_N	101
#define MCT_N2	102
#define MCT_HS	201
#define MCT_HS2	202
