import math

from loadXLS import *
from pramdb import *


def risk(likelihood,impact):
    r=likelihood*impact
    risk=""
    if r<=1:
        risk="Negligible"
    elif r<=4: # Watch out. Ibd risk matrix is not simmetrical
        risk="Low"
    elif r<=9:
        risk="Medium"
    else:
        risk="High"
    return (r,risk)

# A loadXls
# DB pramdb
# Reads controls and applicability from Excel and loads it in the DB
def load_controls(A,DB):
    A.ws = A.book['Assessment']
    e=A.table_to_list('TblThreatEvents')
    events=[i['Threat Events'] for i in e]

    A.ws=A.book['Controls']
    controls = A.table_to_list('TblControls')

    for control in controls:
        ctrlid=control['Control ID']
        origid = control['Original Control ID']
        title = control['Control Title']
        descr = control['Control Description']
        standard = control['Standard']
        NISTfunction=control['NIST Function']
        typeofcontrol = control['Type of Control']
        FR = control['FR']
        FRtitle = control['FR Title']
        if control['Likelihood']:
            like=True
        else:
            like=False
        if control['Impact']:
            impact=True
        else:
            impact=False
        # DB.add_control(ctrlid,origid,title,descr,like,impact,False)
        # def add_control(self, *, ctrlid, standard, nistfunction, fr, frtitle, origid, title, descr, like, impact,commit=True):

        DB.add_control(ctrlid=ctrlid,standard=standard,nistfunction=NISTfunction,typeofcontrol=typeofcontrol,fr=FR,frtitle=FRtitle,origid=origid,
                       title=title,descr=descr,like=like,impact=impact,commit=False)

        for e in events:
            if control[e]:
                DB.set_ctrl_event(ctrlid,e,False)
        DB.commit()

def Reduction_factor(eff):
    f=math.pow(eff,1)*math.exp(eff-1)
    return 1-f

def Results_table(DB,A):
    titles=("Threat","Threat Level", "Asset Id","Asset Name","Impact Type","Impact Level","Standard","NIST Function",
            "Type of Control","FR Title","Ctrl Id","Ctrl Name","ASL","Likelihood","Impact")

    A.create_sheet("Results")
    StartRow=2
    StartCol=2

    A.add_vector(StartRow,StartCol,titles,"Results")

    scenarios=DB.id_scenario(A.assetName)
    row=StartRow
    for s in scenarios:
        S=DB.scenario(s)
        #R=DB.scenario_effectiveness(s)
        ctrls = DB.scenario_applicable_controls(s)
        Asset=DB.asset(S['AssetId'])
        AssetId=DB.id_asset(Asset['Name'])
        asls=DB.asls(AssetId)
        Event=DB.event(S['EventId'])
        EventName=Event['Name']
        ThreatLevel=S['ThreatLevel']
        Impacts=DB.impacts(AssetId)
        AssetName = Asset['Name']
        for i in Impacts:
            ImpactType=i['ImpactType']
            ImpactName=DB.name_impact_type(ImpactType)
            ImpactLevel=i['ImpactLevel']
            for a in asls:
                cid=a['ControlId']
                if cid in ctrls:
                    control=DB.control(cid)
                    standard=control['Standard']
                    NISTfunction = control['NISTFunction']
                    typeofcontrol = control['TypeOfControl']
                    FRtitle = control['FRTitle']
                    controlname= control['Title']
                    applikelihood = control['Likelihood']
                    appimpact = control['Impact']
                    asl=a['ASL']

                    v=(EventName, ThreatLevel, AssetId,AssetName,ImpactName,ImpactLevel,standard,NISTfunction,
                       typeofcontrol,FRtitle,cid,controlname,asl,applikelihood,appimpact)
                    row+=1
                    A.add_vector(row,StartCol,v,"Results")

        # StartColLetter=get_column_letter(StartCol)
        # EndColLetter=get_column_letter(StartCol+len(titles)-1)
        # range=StartColLetter+str(StartRow)+":"+EndColLetter+str(row)
        #
        # try:
        #     del A.ws.tables["data"]
        # except:
        #     pass
        # tab=Table(displayName="data",name="data",ref=range,tableColumns=titles)
        # style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
        #                        showLastColumn=False, showRowStripes=True, showColumnStripes=True)
        # tab.tableStyleInfo = style
        # A.ws.add_table(tab)

        # tab=Table(displayName="XXX",ref="B2:C4")
        # A.ws.add_table(tab)
        A.book.save("assessment.xlsx")




if __name__ == '__main__':
    DB=Prams()
    DB.initialize()
    A=AssessmentXLS()

    # load_controls(A,DB)
    # exit()

    # Load assessment data in DB
    DB.add_asset(A.assetName,A.assetType)
    for impact in A.impacts:
        DB.set_impact(A.assetName,impact['Impact Type'],impact['Level'])

    for asl in A.asls:
        DB.set_asl(A.assetName,asl['Ctrl ID'],asl['ASL'])

    for scenario in A.scenarios:
        DB.create_scenario(A.assetName,scenario['Threat Level'],scenario['Threat Event'])

    #----------------------------------------------

    Results_table(DB,A)

    # scenarios=DB.id_scenario(A.assetName)
    # for s in scenarios:
    #     S=DB.scenario(s)
    #     AssetName=DB.asset(S['AssetId'])
    #     EventName=DB.event(S['EventId'])
    #     ThreatLevel=S['ThreatLevel']
    #
    #     R=DB.scenario_effectiveness(s)
    #
    #     # Gets TSL as the maximum potential impact for the asset
    #     TSL=max([impact['Level'] for impact in A.impacts])
    #     # Identifies the ids of the impact categories with max impact
    #     CriCatIDs=[impact['Impact Type'] for impact in A.impacts if impact['Level']==TSL ]
    #
    #     ELikelihood=0.0
    #     EImpact =0.0
    #
    #     for i in range(TSL,5):
    #         ELikelihood += R['Effectiveness'][i]['Likelihood']
    #         EImpact += R['Effectiveness'][i]['Impact']
    #
    #     # Calculate how many ctrls are applicable to likelihood and impact
    #
    #     NumLikelihood=sum([len(a) for a in R['Controls']['Likelihood']])
    #     NumImpact = sum([len(a) for a in R['Controls']['Impact']])
    #     ELikelihood=ELikelihood/NumLikelihood
    #     EImpact = EImpact / NumImpact
    #
    #
    #     print("Max TSL: ",TSL," in impact categories",CriCatIDs)
    #     print("Control Efficiency for Prevention: ",f"{ELikelihood:.0%}",", for Recovery:",f"{EImpact:.0%}")
    #     redFLikelihood=Reduction_factor(ELikelihood)
    #     redFImpact=Reduction_factor(EImpact)
    #
    #     NTLevel=ThreatLevel*redFLikelihood
    #     print("Initial Threat Strength: ", ThreatLevel, " Updated Threat Strength:", NTLevel)
    #
    #     IImpact=max([l['Level'] for l in A.impacts])
    #     NImpact=IImpact*redFImpact
    #     print ("Initial Impact: ",IImpact," Updated Impact:",NImpact)
    #
    #     risk=risk(NTLevel,NImpact)
    #     print ("Risk is ",risk[1]," (",risk[0],")")
    #     print ("The ineffective controls causing risk are: ",R['Controls']['Ineffective'])





