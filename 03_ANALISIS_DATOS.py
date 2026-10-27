#!/usr/bin/env python3
"""SpringLog-IA - EDA reproducible para logs sintéticos de payment-service."""
from pathlib import Path
import re, json, argparse
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

_FILE = Path(__file__).resolve()
ROOT = _FILE.parent.parent if _FILE.parent.name == 'src' else _FILE.parent
DATA_DIR = ROOT / 'data'
OUT_DIR = ROOT / 'outputs'
CHART_DIR = ROOT / 'charts'
REQUIRED = ['timestamp','app_name','instance_id','severity','thread','logger_class','exception_type','log_message']
PALETTE = {'INFO':'#0F8B8D','WARN':'#E0A126','ERROR':'#D95D39','CRITICAL':'#9B2226'}

# Deliberately conservative masking: logs are sanitized before model feature extraction.
EMAIL_RE = re.compile(r'(?i)\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b')
JWT_RE = re.compile(r'(?i)\bBearer\s+[A-Za-z0-9._~+/=-]{12,}')
KEYVAL_RE = re.compile(r'(?i)\b(password|passwd|pass|secret|api[_-]?key|token)\s*=\s*([^\s,;]+)')
USERID_RE = re.compile(r'(?i)\b(userId|customerId|accountId)\s*=\s*([^\s,;]+)')
CARD_RE = re.compile(r'(?i)\b(card|pan)\s*=\s*([^\s,;]+)')
SEVERITY_RE = re.compile(r'\b(INFO|WARN|ERROR|CRITICAL|DEBUG|FATAL)\b', re.I)

def sanitize_message(value: str) -> str:
    s = str(value)
    s = JWT_RE.sub('Bearer [REDACTED_JWT]', s)
    s = KEYVAL_RE.sub(lambda m: f'{m.group(1)}=[REDACTED]', s)
    s = EMAIL_RE.sub('[REDACTED_EMAIL]', s)
    s = USERID_RE.sub(lambda m: f'{m.group(1)}=[REDACTED_ID]', s)
    s = CARD_RE.sub(lambda m: f'{m.group(1)}=[REDACTED_PAN]', s)
    return s

def style_axes(ax, title, ylabel=None):
    ax.set_title(title, loc='left', fontsize=12, fontweight='bold', color='#16324F', pad=12)
    if ylabel: ax.set_ylabel(ylabel, color='#4B5563')
    ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
    ax.grid(axis='y', alpha=.22); ax.set_axisbelow(True)

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--input', default=str(DATA_DIR / 'payment_service_logs_500.csv'))
    parser.add_argument('--out-dir', default=str(OUT_DIR))
    args = parser.parse_args()
    out = Path(args.out_dir); out.mkdir(parents=True, exist_ok=True)
    charts = CHART_DIR; charts.mkdir(parents=True, exist_ok=True)
    raw = pd.read_csv(args.input, dtype=str, keep_default_na=False)
    missing_cols = sorted(set(REQUIRED) - set(raw.columns))
    if missing_cols: raise ValueError(f'Faltan columnas obligatorias: {missing_cols}')
    df = raw[REQUIRED].copy()
    if len(df) != 500: raise ValueError(f'Se esperaban 500 registros y se encontraron {len(df)}')
    df['timestamp'] = pd.to_datetime(df['timestamp'], errors='raise')
    # Literal None is a sentinel in source, not a CSV null: normalise to explicit category.
    sentinel_none = int(df['exception_type'].str.strip().str.lower().isin(['none','null','']).sum())
    df['exception_type'] = df['exception_type'].replace({'None':'NO_APLICA','none':'NO_APLICA','':'NO_APLICA','null':'NO_APLICA'})
    df['log_message'] = df['log_message'].map(sanitize_message)
    df['date'] = df['timestamp'].dt.strftime('%Y-%m-%d')
    df['hour'] = df['timestamp'].dt.hour
    df['day_name'] = df['timestamp'].dt.day_name()
    df['timestamp'] = df['timestamp'].dt.strftime('%Y-%m-%d %H:%M:%S')
    # Dataset integrity checks.
    duplicates = int(df.duplicated(subset=REQUIRED).sum())
    allowed = {'INFO','WARN','ERROR','CRITICAL'}
    unexpected = sorted(set(df['severity']) - allowed)
    if unexpected: raise ValueError(f'Severidades no esperadas: {unexpected}')
    null_counts = df.replace('',np.nan).isna().sum().to_dict()
    if any(null_counts.values()): raise ValueError(f'Hay faltantes reales después de limpieza: {null_counts}')
    clean_path = out / 'payment_service_logs_500_clean.csv'
    df.to_csv(clean_path, index=False, encoding='utf-8')
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    df.to_csv(DATA_DIR / 'payment_service_logs_500_clean.csv', index=False, encoding='utf-8')
    sev_order=['INFO','WARN','ERROR','CRITICAL']
    severity_counts=df['severity'].value_counts().reindex(sev_order,fill_value=0)
    severity_pct=(severity_counts/len(df)*100).round(2)
    node_table=pd.crosstab(df['instance_id'],df['severity']).reindex(columns=sev_order,fill_value=0)
    node_table['TOTAL']=node_table.sum(axis=1)
    node_table['PORCENTAJE']=node_table['TOTAL']/len(df)*100
    exception_counts=df.loc[df['exception_type']!='NO_APLICA','exception_type'].value_counts().head(10)
    hourly=df.assign(hour_bucket=pd.to_datetime(df['timestamp']).dt.floor('h')).groupby('hour_bucket').size()
    # 1) Severity distribution
    fig,ax=plt.subplots(figsize=(8.4,4.6)); bars=ax.bar(sev_order,severity_counts.values,color=[PALETTE[s] for s in sev_order],width=.62)
    for b,v in zip(bars,severity_counts.values): ax.text(b.get_x()+b.get_width()/2,b.get_height()+3,f'{v} ({v/len(df):.0%})',ha='center',fontsize=9,color='#16324F')
    style_axes(ax,'Distribución de severidad','Registros'); ax.set_ylim(0,max(severity_counts.values)*1.22)
    fig.tight_layout(); fig.savefig(charts/'eda_severidad.png',dpi=180,bbox_inches='tight'); plt.close(fig)
    # 2) Severity by instance stacked bar
    fig,ax=plt.subplots(figsize=(9,4.8)); bottom=np.zeros(len(node_table))
    for s in sev_order:
        vals=node_table[s].values
        ax.bar(node_table.index,vals,bottom=bottom,label=s,color=PALETTE[s]); bottom += vals
    style_axes(ax,'Registros por instancia y severidad','Registros'); ax.tick_params(axis='x',rotation=12); ax.legend(ncol=4,frameon=False,loc='upper left')
    fig.tight_layout(); fig.savefig(charts/'eda_instancias.png',dpi=180,bbox_inches='tight'); plt.close(fig)
    # 3) Hourly volume across whole collection window
    fig,ax=plt.subplots(figsize=(9,4.4)); ax.plot(pd.to_datetime(hourly.index),hourly.values,color='#0F8B8D',linewidth=1.8)
    style_axes(ax,'Volumen de logs por hora','Registros por hora'); ax.tick_params(axis='x',rotation=25)
    fig.tight_layout(); fig.savefig(charts/'eda_hora.png',dpi=180,bbox_inches='tight'); plt.close(fig)
    # 4) Top exception types
    fig,ax=plt.subplots(figsize=(8.5,4.8)); ex=exception_counts.sort_values()
    ax.barh(ex.index,ex.values,color='#2C6E8F'); style_axes(ax,'Tipos de excepción más frecuentes','Registros')
    fig.tight_layout(); fig.savefig(charts/'eda_excepciones.png',dpi=180,bbox_inches='tight'); plt.close(fig)
    # 5) message length by severity
    plot_df=df.assign(message_length=df['log_message'].str.len())
    vals=[plot_df.loc[plot_df.severity==s,'message_length'].values for s in sev_order]
    fig,ax=plt.subplots(figsize=(8.5,4.6)); bp=ax.boxplot(vals,tick_labels=sev_order,patch_artist=True,showfliers=False)
    for patch,s in zip(bp['boxes'],sev_order): patch.set_facecolor(PALETTE[s]); patch.set_alpha(.75)
    style_axes(ax,'Longitud del mensaje por severidad','Caracteres')
    fig.tight_layout(); fig.savefig(charts/'eda_longitud.png',dpi=180,bbox_inches='tight'); plt.close(fig)
    # 6) severity by day
    day_order=list(dict.fromkeys(pd.to_datetime(df['timestamp']).dt.strftime('%Y-%m-%d')))
    daytab=pd.crosstab(pd.to_datetime(df['timestamp']).dt.strftime('%Y-%m-%d'),df['severity']).reindex(index=day_order,columns=sev_order,fill_value=0)
    fig,ax=plt.subplots(figsize=(8.8,4.6)); bottom=np.zeros(len(daytab))
    for s in sev_order:
        vals=daytab[s].values; ax.bar(daytab.index,vals,bottom=bottom,label=s,color=PALETTE[s]); bottom+=vals
    style_axes(ax,'Registros por día y severidad','Registros'); ax.legend(ncol=4,frameon=False); ax.tick_params(axis='x',rotation=0)
    fig.tight_layout(); fig.savefig(charts/'eda_dia.png',dpi=180,bbox_inches='tight'); plt.close(fig)
    stats={
        'n_rows':int(len(df)),'raw_columns':int(len(REQUIRED)),'clean_columns':int(len(df.columns)),
        'source_exception_none_sentinel':sentinel_none,'exact_duplicates':duplicates,'nulls_after_cleaning':{k:int(v) for k,v in null_counts.items()},
        'timestamp_start':str(df.timestamp.min()),'timestamp_end':str(df.timestamp.max()),
        'severity_counts':{k:int(v) for k,v in severity_counts.items()},'severity_percent':{k:float(v) for k,v in severity_pct.items()},
        'instances':{k:int(v) for k,v in df.instance_id.value_counts().items()},
        'instance_severity':{idx:{col:(float(val) if col=='PORCENTAJE' else int(val)) for col,val in row.items()} for idx,row in node_table.to_dict(orient='index').items()},
        'top_exceptions':{k:int(v) for k,v in exception_counts.items()},
        'mean_message_chars':float(df.log_message.str.len().mean()),
        'caveat':'Dataset sintético; los resultados descriptivos no son una estimación de incidentes reales de producción.'
    }
    (out/'eda_summary.json').write_text(json.dumps(stats,indent=2,ensure_ascii=False),encoding='utf-8')
    severity_counts.rename('registros').to_csv(out/'severity_distribution.csv',header=True)
    node_table.to_csv(out/'severity_by_instance.csv')
    exception_counts.rename('registros').to_csv(out/'top_exceptions.csv',header=True)
    print(json.dumps(stats,indent=2,ensure_ascii=False))
    print(f'Clean CSV: {clean_path}')

if __name__=='__main__': main()
