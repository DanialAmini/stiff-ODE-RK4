Attribute VB_Name = "Module1"
Const lambda = -3

Public Function fphi(x)
    fphi = Atn(2 * x)
End Function

Public Function fphi_p(x)
    fphi_p = 2 / (1 + 4 * x ^ 2)
End Function

Public Function func1_(x, y)

    func1_ = lambda * (y - fphi(x)) + fphi_p(x)

End Function

Public Function func2_(x)

    func2_ = Atn(2 * x)
End Function

Public Function func_cp(x, y, dx)
    x1 = x
    y1 = y
    k1 = func1_(x1, y1)
    
    x2 = x + dx
    y2 = y + k1 * dx
    k2 = func1_(x2, y2)
    
    ynew = y + (k1 + k2) * dx / 2
    
    func_cp = ynew

End Function
Public Function func_CN(x, y, dx)

    x1 = x
    y1 = y
    k1 = func1_(x1, y1)
    
    x2 = x + dx
    phi2 = fphi(x2)
    phi_p2 = fphi_p(x2)
    
    ynew = (y / dx + 1 / 2 * (-lambda * phi2 + phi_p2 + k1)) / (1 / dx - lambda / 2)
    
    func_CN = ynew
    
End Function



Public Function func_rk4_(x, y, dx)
    
    x1 = x
    y1 = y
    k1 = func1_(x1, y1)
    
    x2 = x + dx / 2
    y2 = y + k1 * dx / 2
    k2 = func1_(x2, y2)

    x3 = x + dx / 2
    y3 = y + k2 * dx / 2
    k3 = func1_(x3, y3)
    
    x4 = x + dx
    y4 = y + k3 * dx
    k4 = func1_(x4, y4)
    
    ynew = y + (k1 / 6 + k2 / 3 + k3 / 3 + k4 / 6) * dx
    
    func_rk4_ = ynew
    
End Function
