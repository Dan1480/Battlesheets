from oauth2client import service_account
from apiclient import discovery
from pprint import PrettyPrinter
import time
import random as r
import inspect
from gameboard import Gameboard
from coord import Coord
from constants import GAME_WIDTH, GAME_HEIGHT, NOT_HIT, \
HIT, SUNK, MISS, SHIP_LENGTHS
from pprint import PrettyPrinter

#some Google sheets related constants
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = '.\sheets_controller_secret.json'
LOBBYSHEET = 'Players'
#we'll change this later

class SheetsManager():
    def __init__(self):

        credentials = service_account.ServiceAccountCredentials.from_json_keyfile_name(
                      SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        self.service = discovery.build('sheets', 'v4', credentials = credentials)
        self.spreadsheetId = '1wwf3l6GjBeInXSLFLcf71u6DB5CTisCA_dhRqWPGayc'
        self.players_sheet_id = -1 #change later
        self.player_name =''
        self.player_row_in_lobby = -1
        self.player_range = ''
        self.player_status_cell = ''
        self.player_ships_range = ''
        self.opponent_name = ''
        self.opponent_row_in_lobby = -1
        self.opponent_range = ''
        self.opponent_status_cell = ''
        self.opponent_ships_range = ''
        self.game_sheet_name =''
        self.game_sheet_id = -1


    def can_add_player(self, name):
        #See who's currently in the lobby
        values = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheetId,
                range=LOBBYSHEET+'!A2:C').execute().get('values', [])
        if len(values) == 0 \
        or name not in [item[0] for item in values] :
            return True
        else : return False

    def add_player(self, name):
        request_body  = {}
        request_body['values'] = [[name,'Waiting']]
        self.service.spreadsheets().values().\
        append(spreadsheetId=self.spreadsheetId,\
        range=LOBBYSHEET+'!A2:C', valueInputOption = 'RAW', \
        insertDataOption='INSERT_ROWS', body = request_body).execute()

    def put_in_game(self, name):
        self.player_name = name
        current_status = 'Waiting'
        #get current player list
        values = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheetId,
                range=LOBBYSHEET+'!A2:C').execute().get('values', [])

        #there should never be no one in the lobby, but just in case...
        if len(values) == 0:
            return None

        #get my row
        my_row = [i+2 for i,x in enumerate(values) if x[0] == name][0]
        self.player_row_in_lobby = my_row

        while current_status != 'In game':

            new_opponent_row = -1
            new_opponent_name = ''

            #get requested opponent, if player is  requesting someone
            if len(values[my_row -2]) == 3:
                requested_opponent = values[my_row-2][2]
            else : requested_opponent = ''

            rejected = [] #people that want to play, but I reject
            data = []  #to send as batch update to Google sheets API
            body = {}  #also to send as batch update to Google sheets API

            #check if anyone is requesting this player
            for i, player in enumerate(values):
                if player[1] == 'Requesting' and len(player) == 3 \
                and player[2] == name:
                #if an opponenent hasn't already been found
                    if new_opponent_row == -1:
                        new_opponent_row = i + 2
                        new_opponent_name = player[0]
                        self.opponent_row_in_lobby = i + 2
                        self.opponent_name = player[0]
                    #already found an opponent, so add these guys to rejected[]
                    elif player[0] != self.opponent_name:
                        rejected.append(i+2)

            #someone was requesting this player
            if new_opponent_row != -1 :
                #update opponent's status
                data.append(
                                {
                                    "range" : f"{LOBBYSHEET}!B{new_opponent_row}",
                                    "majorDimension" : "ROWS",
                                    "values" : [["In game"]]
                                }
                )
                #update my status
                data.append(
                                {
                                    "range" : f"{LOBBYSHEET}!B{my_row}:C{my_row}",
                                    "majorDimension" : "ROWS",
                                    "values" : [["In game", new_opponent_name]]
                                }
                )
                #find players who requesting my opponent
                more_rejects = [i + 2 for i, player in enumerate(values)
                                if player[1] == 'Requesting'
                                and player[2] == new_opponent_name
                                and player[0] != self.player_name]

                #update status of all rejects
                for row in rejected + more_rejects:
                    data.append(
                                    {
                                        "range": f"{LOBBYSHEET}!B{row}:C{row}",
                                        "majorDimension" : "ROWS",
                                        "values" : [["Waiting",'']]
                                    }
                    )
                body["valueInputOption"]="USER_ENTERED"
                body["data"]=data
                self.service.spreadsheets().values().batchUpdate(
                spreadsheetId=self.spreadsheetId, body=body).execute()
                return new_opponent_name
            #there were no requests for this player but
            #if someone is available, request them
            for player in values :
                if player[1] == 'Waiting' and player[0] != name:
                    #request the first waiting player
                    data.append([
                                {
                                    "range": f"{LOBBYSHEET}!B{my_row}:C{my_row}",
                                    "majorDimension" : "ROWS",
                                    "values" : [["Requesting",player[0]]]
                                }
                    ])
                    body = {
                        "valueInputOption" : "USER_ENTERED",
                        "data" : data
                    }
                    #if this not the person I'm already requesting, request
                    if requested_opponent != player[0]:
                        self.service.spreadsheets().values().batchUpdate(
                        spreadsheetId=self.spreadsheetId, body=body).execute()
                    current_status = 'Requesting'
                    opponent_name = player[0]
                    time.sleep(r.randint(3,5))
                    break
            #Nobody is waiting for a game?  OK, I will wait
            print("Waiting...")
            time.sleep(r.randint(3,5))
            values = self.service.spreadsheets().values().get(
                spreadsheetId=self.spreadsheetId,
                range=f'{LOBBYSHEET}!A2:C').execute().get('values', [])

            #get my row
            my_row = [i+2 for i,x in enumerate(values) if x[0] == name][0]
            self.player_row_in_lobby = my_row
            #get my current status
            current_status = values[my_row - 2][1]

        #someone has put me in a game
        self.opponent_name =  values[my_row - 2][2]
        self.opponent_row_in_lobby = [i+2 for i,x in enumerate(values) if x[0] == self.opponent_name][0]
        return values[my_row - 2][2]

    def setup_game_sheet(self):
        #The player higher up in the lobby creates the sheet
        if self.player_row_in_lobby < self.opponent_row_in_lobby:
            print("Creating game sheet") #debug
            self.game_sheet_name = self.player_name+" vs. "+self.opponent_name
            self.player_range = 'A1:J10'
            self.opponent_range = 'A11:J20'
            self.player_status_cell = 'K1'
            self.opponent_status_cell = 'K11'
            self.player_ships_range = 'K2:L6'
            self.opponent_ships_range = 'K12:L16'
            #create sheet
            body =  {
                      "requests": [
                        {
                          "addSheet": {
                            "properties": {
                              "title": self.game_sheet_name,
                              "gridProperties": {
                                "rowCount": 25,
                                "columnCount": 12
                              },
                              "tabColor": {
                                "red": 0.34,
                                "green": 0.54,
                                "blue": 0.44
                              }
                            }
                          }
                        }
                      ]
                    }
            response = self.service.spreadsheets().\
            batchUpdate(spreadsheetId=self.spreadsheetId,\
            body=body).execute()

            #now set sheetId
            self.game_sheet_id = response['replies'][0]['addSheet']\
            ['properties']['sheetId']
            # set column width
            body = {
              "requests": [
                {
                  "updateDimensionProperties": {
                    "range": {
                      "sheetId": self.game_sheet_id,
                      "dimension": "COLUMNS",
                      "startIndex": 0,
                      "endIndex": 10
                    },
                    "properties": {
                      "pixelSize": 30
                    },
                    "fields": "pixelSize"
                  }
                }
              ]
            }
            self.service.spreadsheets().\
            batchUpdate(spreadsheetId=self.spreadsheetId,body=body).\
            execute()
            #set player statuses
            data = []
            body = {}
            data.append([
                        {
                            "range": f"{self.game_sheet_name}!"+\
                                     f"{self.player_status_cell}",
                            "majorDimension" : "ROWS",
                            "values" : [["Not ready"]]
                        }
            ])
            data.append([
                        {
                            "range": f"{self.game_sheet_name}!"+\
                                     f"{self.opponent_status_cell}",
                            "majorDimension" : "ROWS",
                            "values" : [["Not ready"]]
                        }
            ])

            #set ship lists KEEP WORKING FROM HERE
            ship_names = list(SHIP_LENGTHS.keys())
            values = []
            for s in ship_names:
                values.append([s,"Alive"])
            for x in [self.player_ships_range, self.opponent_ships_range] :
                data.append([
                            {
                                "range": f"{self.game_sheet_name}!"+\
                                         f"{x}",
                                "majorDimension" : "ROWS",
                                "values" : values
                            }
                ])
            body = {
                "valueInputOption" : "USER_ENTERED",
                "data" : data
            }
            self.service.spreadsheets().values().batchUpdate(
            spreadsheetId=self.spreadsheetId, body=body).execute()
            return
        #wait for my opponent to create the sheet
        else :
            self.game_sheet_name = self.opponent_name+" vs. "\
            +self.player_name
            self.opponent_range = 'A1:J10'
            self.player_range = 'A11:J20'
            self.opponent_status_cell = 'K1'
            self.player_status_cell = 'K11'
            self.opponent_ships_range = 'K2:L6'
            self.player_ships_range = 'K12:L16'
            sheet_ready = False
            while not sheet_ready:
                #get the list of sheets
                result = self.service.spreadsheets().\
                get(spreadsheetId=self.spreadsheetId,
                includeGridData=False, \
                fields='sheets/properties(title,sheetId)').execute()

                for sheet in result['sheets']:
                    if sheet['properties']['title'] == self.game_sheet_name:
                        self.game_sheet_id = sheet['properties']['sheetId']
                        sheet_ready = True
                        break
                time.sleep(2)
            #debug
            input("Hit Enter to continue")
    def delete_all_game_sheets(self):
        #get all sheet names and ids
        result = self.service.spreadsheets().\
        get(spreadsheetId=self.spreadsheetId,
        includeGridData=False, \
        fields='sheets/properties(title,sheetId)').execute()

        sheets = [sheet['properties']['sheetId'] for sheet in result['sheets'] \
        if sheet['properties']['title'] != LOBBYSHEET]
        requests = []
        body = {}

        for s in sheets:
            requests.append({"deleteSheet":{"sheetId":s}})
        body['requests']=requests
        response = self.service.spreadsheets().batchUpdate(\
        spreadsheetId=self.spreadsheetId,\
        body=body).execute()
        pp.pprint(response)
    def set_ships(self, gameboard):
        data_range = self.game_sheet_name+"!"+self.player_range
        ship_points = gameboard.get_ship_points()
        #Build command to send to Google Sheets
        body = {'values':[]}

        #add the rows one by one
        for r in range(GAME_HEIGHT):
            row = [''] * GAME_WIDTH
            for c in range(GAME_WIDTH):
                if Coord(c,r) in ship_points.keys():
                    row[c] = NOT_HIT
            body['values'].append(row)

        result = self.service.spreadsheets().values().\
        update(spreadsheetId=self.spreadsheetId,range=data_range,\
        valueInputOption='USER_ENTERED', body=body).execute()

        #now set status to Ready
        body = {'values':[['Ready to start']]}
        result = self.service.spreadsheets().values().\
        update(spreadsheetId=self.spreadsheetId, \
        range = self.game_sheet_name+"!"+self.player_status_cell, \
        valueInputOption='USER_ENTERED',\
        body=body).execute()
        input("Hit Enter to continue...")
    def get_opponent_ships(self, enemy_grid, enemy_ships):
        #wait until opponent is ready
        while True:
            #get opponenent status
            values = self.service.spreadsheets().values().get(
                    spreadsheetId=self.spreadsheetId,
                    range=self.game_sheet_name+"!"+self.opponent_status_cell)\
                    .execute().get('values', [])
            if len(values) > 0 and values[0][0] == 'Ready to start':
                #debug
                print("Opponent is ready")
                #get opponents ship points
                enemy_sheets_points = self.service.spreadsheets().values().get(
                   spreadsheetId=self.spreadsheetId,
                   range=self.game_sheet_name+"!"+self.opponent_range)\
                   .execute().get('values', [])

                if len(enemy_sheets_points) == 0:
                    print("Something went wrong, "
                          "could not get enemy ships.")
                else :
                    #fill empty spots in with blanks
                    for i, row in enumerate(enemy_sheets_points):
                        enemy_grid.append(enemy_sheets_points[i] + \
                        [''] * (GAME_WIDTH - len(row)))

                #get enemy ship statuses
                enemy_sheets_ships = self.service.spreadsheets().values().get(
                   spreadsheetId=self.spreadsheetId,
                   range=self.game_sheet_name+"!"+self.opponent_ships_range)\
                   .execute().get('values', [])
                for ship_status in enemy_sheets_ships :
                    enemy_ships[ship_status[0]]=ship_status[1]
                return
            else:
                print("Waiting for "+self.opponent_name+"...")
                time.sleep(3)
    def wait_for_opponent_turn(self):

#test and debug
if __name__ == '__main__':

    try:
        sm = SheetsManager()
    except Exception as e :
        print(e)
    if sm is not None :
        print("Deleting all game sheets")
        sm.delete_all_game_sheets()
        '''sm.player_name = "Marie"
        sm.opponent_name = "Einstein"
        sm.player_row_in_lobby = 1
        sm.opponent_row_in_lobby = 2
        sm.setup_game_sheet()
        '''
    else : print("Could not create sm object")
