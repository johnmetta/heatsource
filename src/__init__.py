from Dieties.IniParamsDiety import IniParams

def opt(name):
    #See if the file (module name) is to be optimized
    return name.split('.')[-1] in IniParams["psyco"]
    # Return false to turn off all optimization
    #return False