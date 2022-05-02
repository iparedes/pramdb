from sqlite3 import connect,Row

class Prams:
    def __init__(self):
        self.conn=connect("pramdb.db")

        self.actorIntent={'Accidental':1, 'Malicious':2}
        self.actorOrigin = {'External': 1, 'Internal': 2}
        self.actorMotivation = {'High': 4, 'Medium': 2, 'Low':1}
        self.actorPrivilege = {'No Access or Public': 1, 'Unprivileged': 2, 'Moderate': 3, 'Significant':4}
        self.actorSkill = {'Low': 1, 'Medium': 2, 'High': 3, 'Very High': 4}
        self.actorResource = {'Low': 1, 'Medium': 2, 'High': 3, 'Very High': 4}
        self.actorThreatIntel = {
            'No Data': 1,
            'Successful attacks in other industry': 0.5,
            'Attempts to the industry without success': 0.75,
            'Some successful attacks in the industry': 1,
            'Multiple successful attacks in the industry':1.25
        }
        self.actorHistory = {
            'No Data': 1,
            'Few unsuccessful attempts to the company': 0.5,
            'Multiple attempts to the company without success': 0.75,
            'Few successful attempts to the company': 1,
            'Multiple successful attacks to the company':1.25
        }

    def __select_sql(self,sql,vars):
        self.conn.row_factory=Row
        cur=self.conn.cursor()
        cur.execute(sql,vars)
        data=[dict(r) for r in cur.fetchall()]
        #rows=cursor.fetchall()
        return data

    def __insert_sql(self,sql,vars,commit=True):
        try:
            self.conn.execute(sql,vars)
            #cursor.fetchall()
            if commit:
                self.conn.commit()
            return True
        except Exception as e:
            print(e)
            print ("Insert error: ",sql)
            return False

    def __delete_table_sql(self,tablename):
        sql="DELETE FROM "+tablename
        self.conn.execute(sql)

    def commit(self):
        self.conn.commit()

    def initialize(self):
        self.__delete_table_sql('Assessment')
        self.__delete_table_sql('Scenarios')
        self.__delete_table_sql('AssetImpact')
        self.__delete_table_sql('Assets')

        # self.__delete_table_sql('Controls')
        # self.__delete_table_sql('ApplCtrlThreatevent')

    def id_asset(self,name):
        #cursor=self.conn.execute("SELECT Id from 'Assets' where Name=?",(name,))
        #rows=cursor.fetchall()
        sql="SELECT Id from 'Assets' where Name=?"
        var=(name,)
        rows=self.__select_sql(sql,var)
        try:
            a=rows[0]['Id']
            return a
        except:
            return 0

    def id_asset_cat(self,cat):
        # cursor=self.conn.execute("SELECT Id from 'AssetCategory' where Category=?",(cat,))
        # rows=cursor.fetchall()
        sql="SELECT Id from 'AssetCategory' where Category=?"
        var=(cat,)
        rows=self.__select_sql(sql,var)
        try:
            a=rows[0]['Id']
            return a
        except:
            return 0

    def id_impact_type(self,typ):
        # cursor=self.conn.execute("SELECT Id from 'ImpactType' where Name=?",(typ,))
        # rows=cursor.fetchall()
        sql="SELECT Id from 'ImpactType' where Name=?"
        var=(typ,)
        rows=self.__select_sql(sql,var)
        try:
            a=rows[0]['Id']
            return a
        except:
            return 0

    # Needs the name of a Threat Actor
    # Returns its Id
    def id_actor(self,actor):
        sql="SELECT Id from 'ThreatActors' where Actor=?"
        var=(actor,)
        rows=self.__select_sql(sql,var)
        try:
            a=rows[0]['Id']
            return a
        except:
            return 0

    # Needs the name of a Threat Event
    # Returns its Id
    def id_threatEvent(self,event):
        sql="SELECT Id from 'ThreatEvents' where Name=?"
        var=(event,)
        rows=self.__select_sql(sql,var)
        try:
            a=rows[0]['Id']
            return a
        except:
            return 0

    # Needs the name of a Threat Level
    # Returns its Id
    def id_threatLevel(self,level):
        sql="SELECT Id from 'ThreatLevels' where ThreatLevel=?"
        var=(level,)
        rows=self.__select_sql(sql,var)
        try:
            a=rows[0]['Id']
            return a
        except:
            return 0

    # returns the name of an impact type
    def name_impact_type(self,id):
        sql="SELECT Name from ImpactType where Id=?"
        var=(id,)
        rows = self.__select_sql(sql, var)
        try:
            return rows[0]
        except:
            return 0

    # Returns ASLs of an asset (ControlID,ASL)
    def asls(self,idasset):
        #idasset = self.id_asset(asset)
        sql="SELECT ControlID,ASL FROM Assessment where AssetId=?"
        var=(idasset,)
        rows=self.__select_sql(sql,var)
        try:
            return rows
        except:
            return 0

    # Needs the id of a Threat Actor
    # Returns the name of the threat actor
    def threat_actor(self,id):
        sql="SELECT * FROM ThreatActors where Id=?"
        var=(id,)
        a=self.__select_sql(sql,var)
        actor=a[0]
        intent=self.actorIntent[actor['Intent']]
        origin=self.actorOrigin[actor['Origin']]
        actor['Motivation']=intent*origin
        priv=self.actorPrivilege[actor['Privilege']]
        res=self.actorResource[actor['Resources']]
        skill=self.actorSkill[actor['Skill']]
        actor['Capability'] = (priv + res + skill)/3
        actor['Priority']=(actor['Motivation']+actor['Capability'])/2
        return(actor)


    # Needs the id (number) of an asset
    # Returns impacts of an asset (ImpactType, ImpactLevel)
    def impacts(self,idasset):
        #idasset = self.id_asset(asset)
        sql="SELECT ImpactType,ImpactLevel FROM AssetImpact where Asset=?"
        var=(idasset,)
        rows=self.__select_sql(sql,var)
        try:
            return rows
        except:
            return 0

    # returns all fields of a controlid
    def control(self,controlid):
        sql = "SELECT * FROM Controls where Id=?"
        var=(controlid,)
        rows=self.__select_sql(sql,var)
        try:
            return rows[0]
        except:
            return 0

    # Needs the name of an asset
    # Returns the list of ids of the scenarios for that assets
    def id_scenario(self,asset):
        idasset=self.id_asset(asset)
        sql="SELECT Id FROM Scenarios where Asset=?"
        var=(idasset,)
        rows=self.__select_sql(sql,var)
        a=[i['Id'] for i in rows ]
        try:
            return a
        except:
            return 0

    def asset(self,assetid):
        sql="SELECT * FROM Assets WHERE Id=?"
        var=(assetid,)
        rows=self.__select_sql(sql,var)
        if rows:
            r=rows[0]
            name=r['Name']
            category=r['Category']
            R={}
            R['Name']=name
            R['Category']=category
            return R
        else:
            print("Asset Id does not exist")
            return None

    # Needs an event id
    # Returns the event data
    def event(self,id):
        sql="SELECT * FROM ThreatEvents WHERE Id=?"
        var=(id,)
        rows=self.__select_sql(sql,var)
        if rows:
            r=rows[0]
            name=r['Name']
            desc=r['Description']
            R={'Name':name,'Description':desc}
            return R
        else:
            print("Event Id does not exist")

    # Needs an scenario id
    # Returns the scenario data
    def scenario(self,id):
        sql="SELECT * FROM Scenarios WHERE Id=?"
        var=(id,)
        rows=self.__select_sql(sql,var)
        if rows:
            r=rows[0]
            assetid=r['Asset']
            ThreatLevel=r['ThreatLevel']
            eventid=r['Event']
            R={'AssetId':assetid,'ThreatLevel':ThreatLevel,'EventId':eventid}
            return R
        else:
            print("Scenario Id does not exist")

    # ctrlid: control code
    # origid: original control code
    # title:  title of the control
    # descr:  description of the control
    # likelihood: Boolean. True if the control is applicable to likelihood reduction
    # impact: Boolean. True if the control is applicable to impact reduction
    def add_control(self,ctrlid,origid,title,descr,like,impact,commit=True):
        sql="INSERT INTO Controls (Id,OriginalId,Title,Description,Likelihood,Impact) VALUES (?,?,?,?,?,?)"
        var=(ctrlid,origid,title,descr,like,impact,)
        self.__insert_sql(sql, var)

    def add_asset(self,name,cat):
        idcat=self.id_asset_cat(cat)
        if (idcat):

            #cursor = self.conn.execute("INSERT INTO Assets (Category,Name) VALUES (?,?)",(cat,name,))
            sql="INSERT INTO Assets (Category,Name) VALUES (?,?)"
            var=(cat,name,)
            self.__insert_sql(sql, var)
        else:
            print("Asset category does not exist")

    def set_impact(self,asset,typ,value):
        idtype=self.id_impact_type(typ)
        idasset=self.id_asset(asset)
        if (idtype and idasset):
            #cursor = self.conn.execute("INSERT INTO AssetImpact (Asset,'Impact Type','Impact Level') VALUES (?,?,?)",(idasset,idtype,value,))
            sql="INSERT INTO AssetImpact (Asset,'ImpactType','ImpactLevel') VALUES (?,?,?)"
            var=(idasset,idtype,value,)
            self.__insert_sql(sql, var)
        else:
            print("Error setting impact value")

    def set_asl(self,asset,ctrl,val):
        idasset = self.id_asset(asset)
        sql="INSERT INTO Assessment (AssetId,ControlId,ASL) VALUES (?,?,?)"
        var=(idasset,ctrl,val)
        self.__insert_sql(sql, var)

    # Needs a control code, a threat event name
    def set_ctrl_event(self,idctrl,event,commit=True):
        idevent=self.id_threatEvent(event)
        sql="INSERT INTO ApplCtrlThreatEvent (Control,ThreatEvent) VALUES (?,?)"
        var=(idctrl,idevent)
        self.__insert_sql(sql, var, commit)


    # Creates a new scenario
    #   asset: name of the asset
    #   threatLevel: text of the threat level
    #   threatEvent: text of a threat event
    def create_scenario(self,asset,threatLevel,threatEvent):
        idasset=self.id_asset(asset)
        idlevel = self.id_threatLevel(threatLevel)
        idevent = self.id_threatEvent(threatEvent)

        sql = "INSERT INTO Scenarios (Asset,ThreatLevel,Event) VALUES (?,?,?)"
        var = (idasset, idlevel, idevent)
        self.__insert_sql(sql, var)

    # def effectiveness(self,asset):
    #     asls=self.asls(asset)
    #     if asls:
    #         impacts=self.impacts(asset)
    #         for impact in impacts:
    #             TSL=impact[1]
    #             impactname=self.name_impact_type(impact[0])
    #             print(impactname,"TSL: ",TSL)
    #             nctrls=0
    #             effctrls=0
    #             for row in asls:
    #                 ASL=row[1]
    #                 nctrls+=1
    #                 if ASL>=TSL:
    #                     effctrls+=1
    #             print("Eff: ",effctrls/nctrls)
    #     else:
    #         print("No assessment data")


    def scenario_applicable_controls(self,id):
        sql = "SELECT * FROM Scenarios where Id=?"
        var = (id,)
        rows = self.__select_sql(sql, var)
        r = rows[0]
        assetid = r['Asset']
        eventid = r['Event']

        # get controls assessed in the asset that are applicable to the threat event
        sql = "SELECT ControlID FROM Assessment WHERE AssetId=? INTERSECT SELECT Control FROM ApplCtrlThreatEvent WHERE ThreatEvent=?"
        var = (assetid, eventid,)

        aux = self.__select_sql(sql, var)
        # appl_ctrls has the list of applicable controls
        appl_ctrlids = [i['ControlId'] for i in aux]
        return appl_ctrlids

    def scenario_effectiveness(self,id):
        sql="SELECT * FROM Scenarios where Id=?"
        var=(id,)
        rows=self.__select_sql(sql,var)
        r=rows[0]
        assetid=r['Asset']
        levelid=r['ThreatLevel']
        eventid=r['Event']


        # get impacts for the asset
        impacts=self.impacts(assetid)
        # Gets TSL as the maximum potential impact for the asset
        TSL=max([impact['ImpactLevel'] for impact in impacts])
        # Identifies the ids of the impact categories with max impact
        CriCatIDs=[impact['ImpactType'] for impact in impacts if impact['ImpactLevel']==TSL ]

        # CONTROL EFFECTIVENESS
        # get controls assessed in the asset that are applicable to the threat event
        sql="SELECT ControlID FROM Assessment WHERE AssetId=? INTERSECT SELECT Control FROM ApplCtrlThreatEvent WHERE ThreatEvent=?"
        var=(assetid,eventid,)

        aux=self.__select_sql(sql,var)
        # appl_ctrls has the list of applicable controls
        appl_ctrlids=[i['ControlId'] for i in aux]
        controls=[]
        for c in appl_ctrlids:
            control=self.control(c)
            #item={control['Id']:{'Likelihood':control['Likelihood'],'Impact':control['Impact']}}
            item = {'ControlId':control['Id'], 'Likelihood': control['Likelihood'], 'Impact': control['Impact']}
            controls.append(item)

        # Num. of controls applicable to likelihood
        #ctrlsAppLikelihood=len([c for c in controls if c['Likelihood']])
        # Num. of controls applicable to Impact
        #ctrlsApplImpact=len([c for c in controls if c['Impact']])

        # reshape the controls list
        #controldic=[{i['ControlId']:{'Likelihood': i['Likelihood'], 'Impact': i['Impact']}} for i in controls ]
        controldic = {i['ControlId']:{'Likelihood': i['Likelihood'], 'Impact': i['Impact']} for i in controls}
        controls=[]
        # In controldic we have the applicable controls along with its applicability to likelihood and/impact reduction
        # each item in control has a control id as key, and a dictionary with the values of likelihood and impact as value

        # ASLs of all the controls assessed in the asset (some of them may not be applicable to the scenario)
        aux=self.asls(assetid)
        # get only the controls applicable to the Threat in this scenario
        appl_aux=[e for e in aux if e['ControlId'] in appl_ctrlids]
        # reshape the dictionary for easier access
        asls={i['ControlId']:i['ASL'] for i in appl_aux}



        # ctrls_asl_* are lists of 5 elements (one per SL) each one with a list of the controls that achieved that security level
        ctrls_asl_likelihood=[[] for i in range(5)]
        ctrls_asl_impact = [[] for i in range(5)]
        for a in controldic:
            asl=asls[a]
            if controldic[a]['Likelihood']:
                ctrls_asl_likelihood[asl].append(a)
            if controldic[a]['Impact']:
                ctrls_asl_impact[asl].append(a)


        eff_ctrls=[]
        ineff_ctrls=[]
        for c in aux:
            if c['ASL']>=TSL:
                eff_ctrls.append(c['ControlId'])
            else:
                ineff_ctrls.append(c['ControlId'])

        Controls={'Likelihood':ctrls_asl_likelihood,'Impact':ctrls_asl_impact,'Effective':eff_ctrls,'Ineffective':ineff_ctrls}

        Effectiveness=[]
        Effectiveness.append({})
        for i in range(1,5):
            effective_likelihood=0
            effective_impact=0
            for j in range(i,5):
                effective_likelihood+=len(ctrls_asl_likelihood[j])
                effective_impact += len(ctrls_asl_impact[j])
            item={'Likelihood':effective_likelihood, 'Impact':effective_impact}
            Effectiveness.append(item)

        ret={'Effectiveness':Effectiveness,'Controls':Controls}
        return ret




# # Press the green button in the gutter to run the script.
# if __name__ == '__main__':
#     P=Prams()
#     # i=P.id_asset('test')
#     # print(i)
#     # i=P.id_cat('Safety Sytem')
#     # print(i)
#     # P.add_asset('test','Safety System')
#     # P.set_impact('test','Financial',4)
#     # P.set_impact('test', 'Safety and Environment', 2)
#     # P.set_impact('test', 'Financial', 3)
#
#     # P.set_asl('test','ICS.IA.1',2)
#     # P.set_asl('test', 'ICS.IA.2', 1)
#     # P.set_asl('test', 'ICS.IA.4', 4)
#     # P.set_asl('test', 'ICS.IA.3', 3)
#     #
#     # P.effectiveness('test')
#
#     # t=P.threat_actor(4)
#     # print(t)
#     # print(t['Priority'])
#
#    # P.create_scenario(4,2,3)
#     P.scenario_risk(1)
#
#
#
#
#
# # See PyCharm help at https://www.jetbrains.com/help/pycharm/
#
#     # conn=connect("d:\pramdb\pramdb.db")
#     # cursor=conn.execute("SELECT * from 'Asset Category'")
#     # for row in cursor:
#     #     print("ID = ",row[0])
#     #     print("CAT = ", row[1])