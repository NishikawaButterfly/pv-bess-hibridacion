import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import matplotlib.path as mpath
import pandas as pd

def crear_diagrama_profesional():
    """Crea un diagrama unifilar profesional con alto rigor técnico"""
    fig, ax = plt.subplots(figsize=(18, 16))
    ax.set_facecolor('white')
    ax.set_xlim(0, 18)
    ax.set_ylim(0, 16)
    ax.axis('off')
    
    # Estilo profesional
    line_style = dict(color='#005288', linewidth=2.5, zorder=1)
    box_style = dict(boxstyle='round,pad=0.5', facecolor='#f0f7ff', edgecolor='#005288', alpha=0.95, linewidth=1.5)
    mt_style = dict(color='#d62728', linewidth=3, linestyle='-')
    
    # ======== TÍTULO ========
    ax.text(9, 15.5, 'DIAGRAMA UNIFILAR - PLANTA HÍBRIDA PV + BESS 5MW', 
            ha='center', va='center', fontsize=18, fontweight='bold', color='#003366')
    ax.text(9, 15, 'Sistema fotovoltaico 5MW + BESS 5MW/20MWh con integración AMIKIT en media tensión', 
            ha='center', va='center', fontsize=14, color='#005288')
    ax.text(9, 14.6, 'Cumplimiento IEC 61850, IEC 62110 (EMF < 100μT), UNE-EN 50549 y MID para medición', 
            ha='center', va='center', fontsize=12, color='#d62728')
    
    # ======== SISTEMA FOTOVOLTAICO ========
    # PV Array
    ax.text(3, 13.5, 'CAMPO FOTOVOLTAICO 5MW', ha='center', va='center', fontsize=12, fontweight='bold', bbox=box_style)
    for i in range(6):
        x = 2 + i * 0.7
        y = 12.5
        # Símbolo de panel solar profesional
        ax.add_patch(patches.Rectangle((x, y), 0.6, 0.3, fill=True, facecolor='#c9e3ed', edgecolor='#333333'))
        ax.plot([x+0.1, x+0.5], [y+0.15, y+0.15], 'k-', linewidth=1)
        ax.plot([x+0.3, x+0.3], [y+0.05, y+0.25], 'k-', linewidth=1)
    
    # String Inverters (5 x 1MW)
    ax.text(3, 12, 'INVERSORES STRING', ha='center', va='center', fontsize=11, fontweight='bold')
    ax.text(3, 11.8, '5 × 1 MW - Conexión AC 690V', ha='center', va='center', fontsize=10)
    for i in range(5):
        x = 2.2 + i * 0.8
        # Símbolo de inversor profesional
        ax.add_patch(patches.Rectangle((x, 11.3), 0.6, 0.6, fill=True, facecolor='#d1e7f0', edgecolor='#005288'))
        # Símbolo de onda
        t = np.linspace(0, 2*np.pi, 20)
        ax.plot(x + 0.1 + 0.4*t/(2*np.pi), 11.6 + 0.1*np.sin(5*t), 'k-', linewidth=1)
    
    # ======== SISTEMA BESS EN CONTENEDORES ========
    ax.text(14, 13.5, 'SISTEMA BESS EN CONTENEDORES', ha='center', va='center', fontsize=12, fontweight='bold', bbox=box_style)
    ax.text(14, 13.2, '4 × 5 MWh (20\') - LiFePO4 - ENVISION EN-5MWh', ha='center', va='center', fontsize=11)
    
    # Contenedores BESS
    for i in range(4):
        x = 12 + i * 1.5
        y = 12
        # Símbolo de contenedor
        ax.add_patch(patches.Rectangle((x, y), 1.2, 0.8, fill=True, facecolor='#e1f0fa', edgecolor='#005288', linewidth=1.5))
        # Puertas
        ax.add_patch(patches.Rectangle((x+0.9, y+0.1), 0.2, 0.6, fill=True, facecolor='#a9d0e8', edgecolor='#005288'))
        # Ventilación
        for j in range(3):
            ax.add_patch(patches.Rectangle((x+0.2+j*0.3, y+0.7), 0.2, 0.05, fill=True, facecolor='#333333'))
        # Etiqueta
        ax.text(x+0.6, y+0.4, "BESS 5MWh", ha='center', va='center', fontsize=9)
    
    # Sistema de Gestión BESS
    ax.text(16.5, 12.5, 'SISTEMA GESTIÓN BESS', ha='center', va='center', fontsize=10, fontweight='bold')
    ax.add_patch(patches.Circle((16.5, 12), 0.4, fill=True, facecolor='#f0f7ff', edgecolor='#005288'))
    ax.text(16.5, 12, "EMS\nBMS", ha='center', va='center', fontsize=9)
    
    # ======== BARRA AC Y PROTECCIONES ========
    # Barra AC principal
    ax.text(9, 11, 'BARRA AC 690V', ha='center', va='center', fontsize=12, fontweight='bold', bbox=box_style)
    ax.add_patch(patches.Rectangle((6, 10.8), 6, 0.15, fill=True, facecolor='#005288'))
    
    # Interruptor General (ACB)
    ax.text(9, 10.2, 'ACB PRINCIPAL', ha='center', va='center', fontsize=11)
    ax.text(9, 10, '6300A - Icu=65kA - IEC 60947-2', ha='center', va='center', fontsize=10)
    ax.add_patch(patches.Rectangle((8.8, 9.7), 0.4, 0.4, fill=True, facecolor='white', edgecolor='#005288'))
    # Símbolo de interruptor
    ax.plot([8.8, 9.0], [9.9, 9.7], 'k-', linewidth=1.5)
    ax.plot([9.2, 9.0], [9.9, 9.7], 'k-', linewidth=1.5)
    
    # Transformadores de corriente (TC)
    ax.text(7, 9.5, 'TC Medición', ha='center', va='center', fontsize=9)
    ax.add_patch(patches.Circle((7, 9.0), 0.2, fill=True, facecolor='white', edgecolor='#333333'))
    ax.text(7, 9.0, "CT", ha='center', va='center', fontsize=8)
    
    ax.text(11, 9.5, 'TC Protección', ha='center', va='center', fontsize=9)
    ax.add_patch(patches.Circle((11, 9.0), 0.2, fill=True, facecolor='white', edgecolor='#333333'))
    ax.text(11, 9.0, "CT", ha='center', va='center', fontsize=8)
    
    # ======== SISTEMA AMIKIT EN MEDIA TENSIÓN ========
    ax.text(9, 8.5, 'SISTEMA AMIKIT - INTEGRACIÓN MEDIA TENSIÓN', ha='center', va='center', fontsize=12, fontweight='bold', bbox=box_style)
    ax.text(9, 8.2, 'AMK-30MV - Interfaz MT/BT - IEC 61850 - Clase 0.2S MID', ha='center', va='center', fontsize=11, color='#d62728')
    
    # Caja AMIKIT
    ax.add_patch(patches.Rectangle((7.5, 7.5), 3, 1.2, fill=True, facecolor='#e1f0fa', edgecolor='#d62728', linewidth=2))
    
    # Componentes internos AMIKIT (detallados)
    # Celdas GIS
    ax.add_patch(patches.Rectangle((7.7, 7.7), 0.8, 0.5, fill=True, facecolor='#a9d0e8', edgecolor='#333333'))
    ax.text(8.1, 7.95, "Celdas GIS", fontsize=8, ha='center')
    
    # Transformador
    ax.add_patch(patches.Rectangle((8.5, 7.7), 0.8, 0.5, fill=True, facecolor='#a9d0e8', edgecolor='#333333'))
    ax.text(8.9, 7.95, "Trafo", fontsize=8, ha='center')
    
    # PCS
    ax.add_patch(patches.Rectangle((7.7, 8.2), 0.8, 0.5, fill=True, facecolor='#a9d0e8', edgecolor='#333333'))
    ax.text(8.1, 8.45, "PCS", fontsize=8, ha='center')
    
    # Protecciones
    ax.add_patch(patches.Rectangle((8.5, 8.2), 0.8, 0.5, fill=True, facecolor='#a9d0e8', edgecolor='#333333'))
    ax.text(8.9, 8.45, "Protecciones", fontsize=8, ha='center')
    
    # Medidor
    ax.add_patch(patches.Rectangle((8.1, 7.9), 0.8, 0.2, fill=True, facecolor='#ffeb3b', edgecolor='#333333'))
    ax.text(8.5, 8.0, "Medidor 0.2S MID", fontsize=7, ha='center')
    
    # Transformador de tensión (TT)
    ax.text(10.5, 8.5, 'TT Medición', ha='center', va='center', fontsize=9)
    ax.add_patch(patches.Rectangle((10.5, 8.0), 0.2, 0.3, fill=True, facecolor='white', edgecolor='#333333'))
    ax.text(10.5, 8.15, "VT", ha='center', va='center', fontsize=8)
    
    # ======== TRANSFORMADOR HÍBRIDO ========
    ax.text(9, 6.8, 'TRANSFORMADOR HÍBRIDO', ha='center', va='center', fontsize=12, fontweight='bold')
    ax.text(9, 6.5, '5 MVA - 30kV/690V - Dyn11 - Z=6% - EN 50588 - Pérdidas < 5kW', ha='center', va='center', fontsize=11)
    
    # Símbolo de transformador profesional
    ax.add_patch(patches.Circle((9, 6), 0.6, fill=False, edgecolor='#005288', linewidth=2))
    ax.add_patch(patches.Circle((9, 6), 0.4, fill=False, edgecolor='#005288', linewidth=2))
    # Devanados
    for i in range(3):
        angle = i * (2*np.pi/3)
        ax.plot([9 + 0.5*np.cos(angle), 9 + 0.3*np.cos(angle)], 
                [6 + 0.5*np.sin(angle), 6 + 0.3*np.sin(angle)], 'k-', linewidth=1.5)
    
    # ======== PUNTO DE CONEXIÓN Y MEDIDA ========
    ax.text(9, 5, 'PUNTO DE CONEXIÓN A RED 30kV', ha='center', va='center', fontsize=12, fontweight='bold', bbox=box_style)
    ax.text(9, 4.7, 'Límite 5MW - Medida Clase 0.2S - IEC 62053 - MID 2014/32/EU', ha='center', va='center', fontsize=11)
    
    # Símbolo de red eléctrica profesional
    ax.add_patch(patches.Circle((9, 4.2), 0.5, fill=True, facecolor='#fff0f0', edgecolor='#d60000'))
    # Símbolo de onda trifásica
    for i in range(3):
        angle = i * (2*np.pi/3)
        x1 = 9 + 0.3 * np.cos(angle)
        y1 = 4.2 + 0.3 * np.sin(angle)
        x2 = 9 + 0.7 * np.cos(angle)
        y2 = 4.2 + 0.7 * np.sin(angle)
        ax.plot([x1, x2], [y1, y2], 'k-', linewidth=1.5)
    
    # Símbolo de medidor MID
    ax.add_patch(patches.Rectangle((8.5, 3.8), 1, 0.3, fill=True, facecolor='#fffacd', edgecolor='#d6a000'))
    ax.text(9, 4.0, "MEDIDOR ZERA ZMQ 304\nClase 0.2S - MID", ha='center', va='center', fontsize=9)
    
    # Pararrayos
    ax.text(10.5, 5.0, 'PARARRAYOS', ha='center', va='center', fontsize=9)
    ax.add_patch(patches.Rectangle((10.4, 4.8), 0.2, 0.4, fill=True, facecolor='white', edgecolor='#333333'))
    # Símbolo de rayo
    verts = [(10.5, 5.2), (10.45, 5.1), (10.5, 5.15), (10.55, 5.05), (10.5, 5.1)]
    codes = [1, 2, 2, 2, 2]
    path = mpath.Path(verts, codes)
    patch = patches.PathPatch(path, facecolor='none', edgecolor='orange', linewidth=2)
    ax.add_patch(patch)
    
    # ======== SISTEMA DE CONTROL ========
    ax.text(16.5, 10, 'SCADA/EMS CENTRALIZADO', ha='center', va='center', fontsize=12, fontweight='bold', bbox=box_style)
    ax.add_patch(patches.Rectangle((15, 9.3), 3, 1.2, fill=True, facecolor='#f0f7ff', edgecolor='#005288'))
    ax.text(16.5, 9.8, 'Siemens Spectrum Power', ha='center', va='center', fontsize=11)
    ax.text(16.5, 9.5, 'IEC 61850 - Modbus TCP - DNP3 - OPC UA', ha='center', va='center', fontsize=10)
    ax.text(16.5, 9.2, 'Ciberseguridad IEC 62443', ha='center', va='center', fontsize=10, color='#d62728')
    
    # ======== CONEXIONES ========
    # PV a Barra AC
    ax.plot([3, 6], [12.5, 10.8], **line_style)
    ax.plot([3, 6], [11.5, 10.8], **line_style)
    
    # BESS a Barra AC
    ax.plot([14, 6], [12.5, 10.8], **line_style)
    
    # Barra AC a AMIKIT
    ax.plot([9, 9], [10.8, 8.7], **line_style)
    
    # AMIKIT a Transformador
    ax.plot([9, 9], [7.5, 6.6], **line_style)
    
    # Transformador a Red
    ax.plot([9, 9], [5.4, 4.7], **line_style)
    
    # Conexiones de control
    control_style = dict(color='#d60000', linewidth=1.5, linestyle='--', alpha=0.7)
    ax.plot([9, 16.5], [10.8, 9.8], **control_style)  # Barra AC a SCADA
    ax.plot([9, 16.5], [8.7, 9.8], **control_style)    # AMIKIT a SCADA
    ax.plot([9, 16.5], [6, 9.8], **control_style)      # Transformador a SCADA
    ax.plot([9, 16.5], [4.2, 9.8], **control_style)    # Red a SCADA
    
    # ======== LEYENDA Y DETALLES ========
    legend_text = (
        "LEYENDA TÉCNICA:\n"
        "• BESS: Sistema almacenamiento en contenedores 20'\n"
        "• AMIKIT: Integración MT/BT con gestión activa\n"
        "• MID: Sistema de medida certificado para facturación\n"
        "• PCS: Convertidor de potencia bidireccional\n"
        "• ACB: Interruptor automático principal\n"
        "• CT: Transformador de corriente\n"
        "• VT: Transformador de tensión\n"
        "• SCADA/EMS: Supervisión y control centralizado\n"
        "• GIS: Celdas encapsuladas en SF6"
    )
    ax.text(1, 1.5, legend_text, ha='left', va='top', fontsize=10,
            bbox=dict(boxstyle='round,pad=0.5', facecolor='#f5f5f5', edgecolor='#333333'))
    
    ax.text(17, 1.5, f"Rev: 4.0\nFecha: {pd.Timestamp.today().strftime('%d/%m/%Y')}\nCumple IEC 61850/50549/62110", 
            ha='right', va='top', fontsize=9, 
            bbox=dict(boxstyle='round,pad=0.5', facecolor='#f0f0f0', edgecolor='#333333'))
    
    plt.tight_layout()
    plt.savefig('diagrama_unifilar_profesional.png', dpi=300, bbox_inches='tight')
    plt.close()
    return 'diagrama_unifilar_profesional.png'

# Ejecutar la generación del diagrama
crear_diagrama_profesional()
print("Diagrama generado: 'diagrama_unifilar_profesional.png'")