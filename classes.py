#%%
##############################################################################
############################## Element Types #################################
##############################################################################

# Element classses from the excel file
class Element:
    def __init__(self, internal_bus_location, manager, owner, type_of_contract):
        self.internal_bus_location = internal_bus_location
        self.manager = manager
        self.owner = owner
        self.type_of_contract = type_of_contract

# class Load extends Element class and adds the values of: load, charge type, power contracted, power factor.
class Load(Element):
    def __init__(self, internal_bus_location, manager, owner, type_of_contract, load, charge_type, power_contracted, power_factor):
        super().__init__(internal_bus_location, manager, owner, type_of_contract)
        self.load = load
        self.charge_type = charge_type
        self.power_contracted = power_contracted
        self.power_factor = power_factor
# class Generator extends Element class and adds the values of: type_of_generator, P_max_kW, Q_max_kW, Q_min_kW.
class Generator(Element):
    def __init__(self, internal_bus_location, manager, owner, type_of_contract, type_of_generator, P_max_kW, Q_max_kW, Q_min_kW):
        super().__init__(internal_bus_location, manager, owner, type_of_contract)
        self.type_of_generator = type_of_generator
        self.P_max_kW = P_max_kW
        self.Q_max_kW = Q_max_kW
        self.Q_min_kW = Q_min_kW
# class Storage extends Element class and adds the values of: battery_type, energy_capacity_kVAh, energy_min_percent, charge_efficiency_percent, discharge_efficiency_percent, initial_state_percent, p_charge_max_kW, p_discharge_max_kW.
class Storage(Element):
    def __init__(self, internal_bus_location, manager, owner, type_of_contract, battery_type, energy_capacity_kVAh, energy_min_percent, charge_efficiency_percent, discharge_efficiency_percent, initial_state_percent, p_charge_max_kW, p_discharge_max_kW):
        super().__init__(internal_bus_location, manager, owner, type_of_contract)
        self.battery_type = battery_type
        self.energy_capacity_kVAh = energy_capacity_kVAh
        self.energy_min_percent = energy_min_percent
        self.charge_efficiency_percent = charge_efficiency_percent
        self.discharge_efficiency_percent = discharge_efficiency_percent
        self.initial_state_percent = initial_state_percent
        self.p_charge_max_kW = p_charge_max_kW
        self.p_discharge_max_kW = p_discharge_max_kW
# class Chargring_Station extends Element class and adds the values of:, p_charge_max_kW, p_discharge_max_kW, charge_efficiency_percent, discharge_efficienc_percent, energy_capacity_max_kWh, place_start, place_end.
class Charging_Station(Element):
    def __init__(self, internal_bus_location, manager, owner, type_of_contract, p_charge_max_kW, p_discharge_max_kW, charge_efficiency_percent, discharge_efficiency_percent, energy_capacity_max_kWh, place_start, place_end):
        super().__init__(internal_bus_location, manager, owner, type_of_contract)
        self.p_charge_max_kW = p_charge_max_kW
        self.p_discharge_max_kW = p_discharge_max_kW
        self.charge_efficiency_percent = charge_efficiency_percent
        self.discharge_efficiency_percent = discharge_efficiency_percent
        self.energy_capacity_max_kWh = energy_capacity_max_kWh
        self.place_start = place_start
        self.place_end = place_end
##############################################################################
############################## Other Elements ################################
##############################################################################
class Vehicle:
    def __init__(self, id, manager, owner, type_of_contract, type_of_vehicle, energy_capacity_max_kWh, p_charge_max_kW, p_discharge_max_kW, charge_efficiency_percent, discharge_efficiency_percent, initial_state_SOC_percent, minimun_technical_SOC_percent):
        self.id = id
        self.manager = manager
        self.owner = owner
        self.type_of_contract = type_of_contract
        self.type_of_vehicle = type_of_vehicle
        self.energy_capacity_max_kWh = energy_capacity_max_kWh
        self.p_charge_max_kW = p_charge_max_kW
        self.p_discharge_max_kW = p_discharge_max_kW
        self.charge_efficiency_percent = charge_efficiency_percent
        self.discharge_efficiency_percent = discharge_efficiency_percent
        self.initial_state_SOC_percent = initial_state_SOC_percent
        self.minimun_technical_SOC_percent = minimun_technical_SOC_percent
    def set_profile(self, arrive_time_period=None, departure_time_period=None,
                    place=None, used_soc_percent_arriving=None, soc_percent_arriving=None,
                    soc_required_percent_exit=None, p_charge_max_constracted_kw=None,
                    p_discharge_max_constracted_kw=None, charge_price=None, discharge_price=None):
        self.arrive_time_period = arrive_time_period
        self.departure_time_period = departure_time_period
        self.place = place
        self.used_soc_percent_arriving = used_soc_percent_arriving
        self.soc_percent_arriving = soc_percent_arriving
        self.soc_required_percent_exit = soc_required_percent_exit
        self.p_charge_max_constracted_kw = p_charge_max_constracted_kw
        self.p_discharge_max_constracted_kw = p_discharge_max_constracted_kw
        self.charge_price = charge_price
        self.discharge_price = discharge_price
# class peers, which containts the properties: type_of_peer, type_of_contract, owner.
class Peer:
    def __init__(self, id, type_of_peer, type_of_contract):
        self.id = id
        self.type_of_peer = type_of_peer
        self.type_of_contract = type_of_contract
    # Property method that sets the buy_price_mu and sell_price_mu
    def set_profile(self, buy_price_mu=None, sell_price_mu=None):
        self.buy_price_mu = buy_price_mu
        self.sell_price_mu = sell_price_mu
##############################################################################
############################## Data: Full Excell #############################
##############################################################################
# class Data constains the properties: simulation_periods, periods_duration_min, objective_functions_list, a dict called network_information, a list of objects of the class Vehicle, a list of objects of the class Load, a list of objects of the class Generator, a list of objects of the class Storage, a list of objects of the class Charging_Station, a list of objects of the class Peer.
class Data:
    def __init__(self, simulation_periods, periods_duration_min, objective_functions_list, network_information, vehicle_list, load_list, generator_list, storage_list, charging_station_list, peer_list):
        self.simulation_periods = simulation_periods
        self.periods_duration_min = periods_duration_min
        self.objective_functions_list = objective_functions_list
        self.network_information = network_information
        self.vehicle_list = vehicle_list
        self.load_list = load_list
        self.generator_list = generator_list
        self.storage_list = storage_list
        self.charging_station_list = charging_station_list
        self.peer_list = peer_list