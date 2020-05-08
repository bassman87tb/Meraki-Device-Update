#-----------Update Meraki Switchport-----------#
#  Bulk updates device information for a       #
#  single organization.  Additionally, it      #
#  updates the switchports of a list of Meraki #
#  switches based on an excel spreadsheet.     #
#----------------------------------------------#
#  Author: Trevor Butler                       #
#  Company: LookingPoint                       #
#  Version: 1.0                                #
#----------------------------------------------#
#  Input Arguments:                            #
#      arg1 Meraki API Key - String            #
#      arg2 Excel File location - String       #
#----------------------------------------------#



# Library Includes
import xlrd 
import sys
from os.path import isfile
from meraki_sdk.meraki_sdk_client import MerakiSdkClient
from meraki_sdk.exceptions.api_exception import APIException



#Global Variables
meraki_api = MerakiSdkClient(str(sys.argv[1]))
File_location = str(sys.argv[2])
switch_ports_controller = meraki_api.switch_ports
devices_controller = meraki_api.devices

#Definitions
	#return workbook based on filename
	#after checking if it exists
def get_excel_workbook_object(fname):
	if not isfile(fname):
		print ('File doesnt exist: ',fname)

    # Open the workbook and return object
	wrkbk = xlrd.open_workbook(fname)
	return wrkbk

	#return the sheet from workbook at index (default = 0)
def get_excel_sheet_object(wrkbk_arg, idx=0):
    sheet = wrkbk_arg.sheet_by_index(idx)
    print ('Retrieved worksheet: %s' % sheet.name)
    return sheet
	
	#Read a row(idx_arg) from sheet (sheet_arg) and return a dictionary 
	#with column header as key, and cell value as the value.
def get_row_object(sheet_arg, hdr_off_arg=0, idx_arg=1):
	row = sheet_arg.row(hdr_off_arg+idx_arg)
	header = sheet_arg.row(hdr_off_arg)
	dic = {}
	for idx, col in enumerate(row):
		dic[str(header[idx].value)] = str(col.value)
	#print('Retrieved row',idx_arg)
	#print(dic)
	return dic	
	
	#Get organization ID based on the organization name
def get_org_id(org_arg):
	try:
		orgs = meraki_api.organizations.get_organizations()
	except APIException as e:
		print(e)
	
	for idx, org in enumerate(orgs):
		if org.get("name") == org_arg:
			return org.get("id")
	print("Organization not found")
	return
	
	#Get list of networks in a single organization represented
	#in a List of Dictionaries
def get_networks(org_id_arg):
	params = {}
	params["organization_id"] = org_id_arg
	
	try:
		nets = meraki_api.networks.get_organization_networks(params)
		return nets
	except APIException as e:
		print(e)
	
	#Return the network ID for a particular network based on name
def get_net_id(nets_arg, net_name_arg):
	for idx, net in enumerate(nets_arg):
		if net.get("name") == net_name_arg:
			return net.get("id")
	print("Network not found")
	return
	
	#Read in the row_object dictionary and update the device with
	#the information using network id and serial number
def update_device(nets_arg, dev_obj_arg):
	collect = {}
	collect['network_id'] = get_net_id(nets_arg, dev_obj_arg.get("network name"))
	collect['serial'] = dev_obj_arg.get("switch serial")
	
	#Device attributes
	update_network_device = {}
	update_network_device['name'] = dev_obj_arg.get("switch name")
	update_network_device['tags'] = dev_obj_arg.get("tags")
	if dev_obj_arg.get("address") != '':
		update_network_device['address'] = dev_obj_arg.get("address")
		update_network_device['moveMapMarker'] = True
	if dev_obj_arg.get("notes") != '':
		update_network_device['notes'] = dev_obj_arg.get("notes")
	
	collect['update_network_device'] = update_network_device
	
	try:
		result = devices_controller.update_network_device(collect)
		print('Completed '+str(dev_obj_arg.get("switch name"))+' update')
	except APIException as e:
		print(e)

	#Read in the row_object dictionary and update the switchport with
	#the information using device serial number.
def update_switchport(sp_obj_arg):
	collect = {}
	collect['serial'] = sp_obj_arg.get("switch serial")
	collect['number'] = int(float(sp_obj_arg.get("switchport number")))
	
	#switchport attributes
	update_device_switch_port = {}
	update_device_switch_port['name'] = sp_obj_arg.get("switchport description")
	update_device_switch_port['tags'] = sp_obj_arg.get("tags")
	update_device_switch_port['enabled'] = True
		#POE enable
	if sp_obj_arg.get("poe enabled") != '':
		if sp_obj_arg.get("poe enabled") == 'yes':
			update_device_switch_port['poeEnabled'] = True
		elif sp_obj_arg.get("poe enabled") == 'no':
			update_device_switch_port['poeEnabled'] = False
		#Port Type
	update_device_switch_port['type'] = sp_obj_arg.get("port type")
	if sp_obj_arg.get("port type") == 'access':
		if sp_obj_arg.get("data vlan") != '':
			update_device_switch_port['vlan'] = int(float(sp_obj_arg.get("data vlan")))
		if sp_obj_arg.get("voice vlan") != '':
			update_device_switch_port['voiceVlan'] = int(float(sp_obj_arg.get("voice vlan")))
	elif sp_obj_arg.get("port type") == 'trunk':
		if sp_obj_arg.get("native vlan") != '':
			update_device_switch_port['vlan'] = int(float(sp_obj_arg.get("native vlan")))
		if sp_obj_arg.get("allowed vlans") != '':
			update_device_switch_port['allowedVlans'] = sp_obj_arg.get("allowed vlans")
		#Rapid Spanning-tree
	if sp_obj_arg.get("rstp enabled") != '':
		if sp_obj_arg.get("rstp enabled") == 'yes':
			update_device_switch_port['rstpEnabled'] = True
		if sp_obj_arg.get("rstp enabled") == 'no':
			update_device_switch_port['rstpEnabled'] = False
		#Spanning-tree Guard
	if sp_obj_arg.get("stp guard") != '':
		if sp_obj_arg.get("stp guard") == 'disabled':
			update_device_switch_port['stpGuard'] = 'disabled'
		elif sp_obj_arg.get("stp guard") == 'root guard':
			update_device_switch_port['stpGuard'] = 'root guard'
		elif sp_obj_arg.get("stp guard") == 'bpdu guard':
			update_device_switch_port['stpGuard'] = 'bpdu guard'
		elif sp_obj_arg.get("stp guard") == 'loop guard':
			update_device_switch_port['stpGuard'] = 'loop guard'
		#UDLD
	if sp_obj_arg.get("udld") != '':
		update_device_switch_port['udld'] = sp_obj_arg.get("udld")

	collect['update_device_switch_port'] = update_device_switch_port
	
	try:
		result = switch_ports_controller.update_device_switch_port(collect)
		print('Completed '+str(sp_obj_arg.get("switch name"))+' switchport ' +str(collect.get("number"))+ ' update')
	except APIException as e:
		print(e)
	
	
#Main Program
if __name__=='__main__':
	# Get Excel workbook used by this script
	workbook = get_excel_workbook_object(File_location)

	# Get Organization and network information
	sheet = get_excel_sheet_object(workbook, 0)
	OrganizationId = get_org_id(get_row_object(sheet, 0, 1).get("organization name"))
	if OrganizationId != 'None':
		NetworkList = get_networks(OrganizationId)
		for row_idx in range(1, sheet.nrows-2):    # Iterate through rows
			update_device(NetworkList, get_row_object(sheet, 2, row_idx))

	print(' ')

	# Open Workbook and update switchports
	for idx in range (1, workbook.nsheets):  # Iterate though switchport worksheets
		sheet = get_excel_sheet_object(workbook, idx)
		for row_idx in range(1, sheet.nrows):    # Iterate through rows
			update_switchport(get_row_object(sheet, 0, row_idx))