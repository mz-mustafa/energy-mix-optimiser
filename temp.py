
    def set_bess_parameters(self, y, m, d, h, starting):

        #assumption that sim starts with full reserve
        #and a status of 0, so charge req is 0
        if y==1 and m == 1 and d == 1 and h == 0 and starting:
            return 0
        
        #extract BESS sources
        bess_sources = [src for src in self.src_list if src.metadata['type']['value'] == 'BESS']
        #iterate over sources
        bess_charging_energy = 0
        for src in bess_sources:

            src_hourly_data = src.ops_data[y]['months'][m]['days'][d]['hours'][h]
            #i -1 and -2, -3, then capacity and reserve =0
        
            if src_hourly_data['status'] in [-1, -2, -3]:

                src_hourly_data['capacity'] = src_hourly_data['reserve'] = 0
            
            else:

                if starting:
                    #check reserve for previous hour.
                    year, month, day, hour = self.previous_hour(y,m,d,h)
                    src_prev_hour_data = src.ops_data[year]['months'][month]['days'][day]['hours'][hour]
                    #if bess has been trickle charging undisturbed
                    if src_prev_hour_data['status'] == 0:
                        
                        #keep trickle charging
                        src_hourly_data['status'] = 0
                        src_hourly_data['reserve'] = src_prev_hour_data['reserve']
                        bess_charging_energy += src_hourly_data['capacity'] * 0.01

                    #remaining status is 1/ discharging and 2 full charging
                    else:
                        #probably not needed because for discharging this will always be true
                        if src_prev_hour_data['reserve'] < src_prev_hour_data['capacity']:
                            
                            #full charge
                            src_hourly_data['status'] = 2
                            max_charge_energy_av = src_prev_hour_data['capacity']/4
                            charging_req = src_prev_hour_data['capacity'] - src_prev_hour_data['reserve']
                            #this will be charging that can happen during the hour.
                            actual_charging = min(charging_req, max_charge_energy_av)
                            #we are assuming that the bess will be charged throughout, which might not be the case.
                            bess_charging_energy += actual_charging
                            
                            #for now, we add 50% of actual charging to the reserve.
                            #to simulate that the full cap may not be achieved.
                            src_hourly_data['reserve'] = min(src_hourly_data['capacity'], src_prev_hour_data['reserve'] + (max_charge_energy_av/2))
                            
                            if src_hourly_data['reserve'] == src_hourly_data['capacity']:
                                src_hourly_data['status'] == 0

                        else:
                            src_hourly_data['status'] = 0
                            src_hourly_data['reserve'] = src_prev_hour_data['reserve']
                #if the function is ending
                else:
                    #this means that the bess is still in full charge
                    if src_hourly_data['status'] == 2:
                        #and that the reserve can be topped up.
                        year, month, day, hour = self.previous_hour(y,m,d,h)
                        src_prev_hour_data = src.ops_data[year]['months'][month]['days'][day]['hours'][hour]

                        src_hourly_data['reserve'] = min(src_hourly_data['capacity'], 
                                                         src_hourly_data['reserve'] + (
                                                             0.5*src_prev_hour_data['capacity']/4))
                        if src_hourly_data['reserve'] == src_hourly_data['capacity']:
                            src_hourly_data['status'] == 0
                    
                    #there is no need to test for 0 (trickle) and 1 (discharging) conditions
                    #in case of 0, there is no change to what is set in the starting block
                    #in case of 1, other simulate functions changed the state
                    #so it means that reserve has already been reduced and status changed.
        return bess_charging_energy

def has_stable_source(self,year):

        for src in self.src_list:
            # Check if 'stability' in metadata and its value is 'STABLE'
            if src.metadata.get('stability', {}).get('value') == 'STABLE':
                # Check if the source's start year is not later than the given year
                if src.config.get('start_year') <= year:
                    return True  # Found at least one stable source meeting the criteria
        return False  # No stable source found meeting the criteria
    
    def scenario_includes_renewable_src(self, year):
        
        for src in self.src_list:
            
            if src.metadata.get('type', {}).get('value') == 'R':
                # Check if the source's start year is not later than the given year
                if src.config.get('start_year') <= year:
                    return True  # Found at least one stable source meeting the criteria
        return False  # No renewable source found meeting the criteria
    
    def scenario_includes_captive_src(self, year):

        for src in self.src_list:
            
            if src.metadata.get('finance', {}).get('value') == 'CAPTIVE':
                # Check if the source's start year is not later than the given year
                if src.config.get('start_year') <= year:
                    return True  # Found at least one stable source meeting the criteria
        return False  # No renewable source found meeting the criteria