import math

from loadXLS import *
from pramdb import *

STARTROW=2
STARTCOL=2

# Rows Likelihood, Cols Impact
RISK_MATRIX=[
    [0,0,0,0,0],
    [0,1,1,1,1],
    [0,1,1,2,3],
    [0,1,2,3,4],
    [0,1,3,4,4]
]

Risk_Names=["N/A","Negligible","Low","Medium","High"]

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
    #A.ws = A.book['Assessment']
    e=A.table_to_list('TblThreatEvents','Assessment')
    events=[i['Threat Events'] for i in e]

    #A.ws=A.book['Controls']
    controls = A.table_to_list('TblControls','Controls')

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

def Results_table(DB, A, startrow=STARTROW,startcol=STARTCOL ):
    titles=("Threat","Threat Level","Asset Id","Asset Name","Impact Type","Impact Level","Standard","NIST Function",
            "Type of Control","FR Title","Ctrl Id","Ctrl Name","ASL","Likelihood","Impact","Gap","Eff Factor",
            "Eff Factor Likelihood","Eff Factor Impact")

    rows=[]

    # A.create_sheet("Results")
    A.delete_table("data")
    #A.book.save("assessment.xlsx")
    #exit()



    ImpactLevelColLetter=get_column_letter(startcol+titles.index("Impact Level"))
    AslColLetter=get_column_letter(startcol+titles.index("ASL"))
    GapColLetter = get_column_letter(startcol + titles.index("Gap"))
    LikColLetter=get_column_letter(startcol + titles.index("Likelihood"))
    ImpColLetter = get_column_letter(startcol + titles.index("Impact"))

    rows.append(titles)
    #A.add_vector(startrow, startcol, titles, "Results")

    scenarios=DB.id_scenario(A.assetName)

    row=startrow

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

                    row+=1
                    Gap="="+ImpactLevelColLetter+str(row)+"-"+AslColLetter+str(row)
                    EffImpact="=IF("+GapColLetter+str(row)+"<=0,1,-"+GapColLetter+str(row)+")"

                    EffFactorLikelihood="=IF("+LikColLetter+str(row)+"=0,0,IF("+GapColLetter+str(row)+"<=0,1,-"+GapColLetter+str(row)+"))"
                    EffFactorImpact = "=IF(" + ImpColLetter + str(row) + "=0,0,IF(" + GapColLetter + str(
                        row) + "<=0,1,-" + GapColLetter + str(row) + "))"

                    v=(EventName, ThreatLevel, AssetId,AssetName,ImpactName,ImpactLevel,standard,NISTfunction,
                     typeofcontrol,FRtitle,cid,controlname,asl,applikelihood,appimpact,Gap,EffImpact,EffFactorLikelihood,
                       EffFactorImpact)
                    #A.add_vector(row,startcol,v,"Results")
                    rows.append(v)

        A.create_table("data",startrow,startcol,rows,"Results")

        nrows=len(rows)
        rows=[]
        titlesSum=("Impact Type","TSL","# applicable controls","App Likelihood","App Impact","Sum Eff Likelihood",
                   "Sum Eff Impact","Eff Likelihood","Eff Impact","Eff Likelihood adj","Eff Impact adj")
        rows.append(titlesSum)

        row=startrow
        startcolsum=startcol+len(titles)+4
        startcolsumLetter=get_column_letter(startcolsum)
        ImpactTypeLetter=get_column_letter(startcol+titles.index("Impact Type"))
        ImpactTypeCol="$"+ImpactTypeLetter+str(startrow+1)+":$"+ImpactTypeLetter+str(startrow+nrows-1)

        AppLikLetter=get_column_letter(startcol+titles.index("Likelihood"))
        AppLikCol="$"+AppLikLetter+str(startrow+1)+":$"+AppLikLetter+str(startrow+nrows-1)

        AppImpLetter = get_column_letter(startcol + titles.index("Impact"))
        AppImpCol="$"+AppImpLetter+str(startrow+1)+":$"+AppImpLetter+str(startrow+nrows-1)

        EffFactorLetter=get_column_letter(startcol + titles.index("Eff Factor"))
        EffFactorCol="$"+EffFactorLetter+str(startrow+1)+":$"+EffFactorLetter+str(startrow+nrows-1)

        SumEffLikLetter=get_column_letter(startcolsum+titlesSum.index("Sum Eff Likelihood"))
        SumEffImpLetter = get_column_letter(startcolsum + titlesSum.index("Sum Eff Impact"))

        AppLikSumLetter=get_column_letter(startcolsum+titlesSum.index("App Likelihood"))
        AppImpSumLetter = get_column_letter(startcolsum + titlesSum.index("App Impact"))

        EffLikSumLetter=get_column_letter(startcolsum+titlesSum.index("Eff Likelihood"))
        EffImpSumLetter = get_column_letter(startcolsum + titlesSum.index("Eff Impact"))

        for i in Impacts:
            row+=1
            ImpactType=i['ImpactType']
            ImpactName = DB.name_impact_type(ImpactType)
            TSL=i['ImpactLevel']
            numAppCtrls="=COUNTIF("+ImpactTypeCol+","+startcolsumLetter+str(row)+")"
            numAppLikelihood="=COUNTIFS("+ImpactTypeCol+"," + startcolsumLetter + str(row) + "," + AppLikCol + ",1)"
            numAppImpact = "=COUNTIFS(" + ImpactTypeCol+"," + startcolsumLetter + str(row) + "," + AppImpCol + ",1)"
            sumEffLik="=SUMIFS("+EffFactorCol+","+ImpactTypeCol+"," + startcolsumLetter + str(row) + "," + AppLikCol + ",1)"
            sumEffImp = "=SUMIFS(" + EffFactorCol + "," + ImpactTypeCol + "," + startcolsumLetter + str(row) + "," + AppImpCol + ",1)"
            effLik="="+SumEffLikLetter+str(row)+"/"+AppLikSumLetter+str(row)
            effImp = "=" + SumEffImpLetter + str(row) + "/" + AppImpSumLetter + str(row)
            effLikAdj="="+AppLikLetter+str(row)+"IF("+EffLikSumLetter+str(row)+"<=0,0,"+EffLikSumLetter+str(row)+")"
            effImpAdj = "="+AppImpLetter+str(row)+"IF(" + EffImpSumLetter + str(row) + "<=0,0," + EffImpSumLetter + str(row) + ")"

            rows.append((ImpactName,TSL,numAppCtrls,numAppLikelihood,numAppImpact,sumEffLik,sumEffImp,effLik,effImp,effLikAdj,effImpAdj))

        A.create_table("summary", startrow, startcolsum, rows, "Results")

        rowAsset=startrow+len(rows)+2
        rowThreat=rowAsset+1

        A.set_cell(rowAsset,startcolsum,"Asset","Results")

        # assetlist=DB.assets()
        # assets=[a['Name'] for a in assetlist]
        # assetstr=",".join(assets)
        # A.add_data_validation(rowAsset,startcolsum+1,assetstr,"Results")

        A.book.save("assessment.xlsx")







if __name__ == '__main__':
    DB=Prams()
    DB.initialize()
    A=AssessmentXLS("assessment.xlsx")

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





