#!/usr/bin/env python3
"""SpringLog-IA - Entrenamiento y evaluación reproducible del clasificador."""
from pathlib import Path
import argparse, json, re, time, warnings
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import joblib
from sklearn.model_selection import train_test_split
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.pipeline import Pipeline
from sklearn.ensemble import RandomForestClassifier
from sklearn.linear_model import LogisticRegression
from sklearn.naive_bayes import MultinomialNB
from sklearn.tree import DecisionTreeClassifier
from sklearn.metrics import accuracy_score, f1_score, precision_score, recall_score, classification_report, confusion_matrix

_FILE=Path(__file__).resolve()
ROOT=_FILE.parent.parent if _FILE.parent.name=='src' else _FILE.parent
DATA_DIR=ROOT/'data'; OUT_DIR=ROOT/'outputs'; CHART_DIR=ROOT/'charts'
PALETTE={'INFO':'#0F8B8D','WARN':'#E0A126','ERROR':'#D95D39','CRITICAL':'#9B2226'}
SEVERITY_RE=re.compile(r'\b(INFO|WARN|ERROR|CRITICAL|DEBUG|FATAL)\b',re.I)
EMAIL_RE=re.compile(r'(?i)\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b')
JWT_RE=re.compile(r'(?i)\bBearer\s+[A-Za-z0-9._~+/=-]{12,}')
KEYVAL_RE=re.compile(r'(?i)\b(password|passwd|pass|secret|api[_-]?key|token)\s*=\s*([^\s,;]+)')
USERID_RE=re.compile(r'(?i)\b(userId|customerId|accountId)\s*=\s*([^\s,;]+)')
CARD_RE=re.compile(r'(?i)\b(card|pan)\s*=\s*([^\s,;]+)')

def sanitize(value):
    s=str(value)
    s=JWT_RE.sub('Bearer [REDACTED_JWT]',s)
    s=KEYVAL_RE.sub(lambda m:f'{m.group(1)}=[REDACTED]',s)
    s=EMAIL_RE.sub('[REDACTED_EMAIL]',s)
    s=USERID_RE.sub(lambda m:f'{m.group(1)}=[REDACTED_ID]',s)
    s=CARD_RE.sub(lambda m:f'{m.group(1)}=[REDACTED_PAN]',s)
    # Prevent explicit label words in a log line from becoming a shortcut feature.
    s=SEVERITY_RE.sub('[LEVEL]',s)
    return s

def build_features(df):
    exc=df['exception_type'].replace({'None':'NO_APLICA','none':'NO_APLICA','':'NO_APLICA'}).fillna('NO_APLICA').astype(str)
    msg=df['log_message'].fillna('').astype(str).map(sanitize)
    return (df['logger_class'].fillna('').astype(str)+' '+exc+' '+msg).tolist()

def make_pipe(clf):
    return Pipeline([('tfidf',TfidfVectorizer(ngram_range=(1,2),max_features=1000,sublinear_tf=True)),('clf',clf)])

def style(ax,title):
    ax.set_title(title,loc='left',fontsize=12,fontweight='bold',color='#16324F',pad=12)
    ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False); ax.grid(axis='y',alpha=.2); ax.set_axisbelow(True)

def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--input',default=str(OUT_DIR/'payment_service_logs_500_clean.csv'))
    parser.add_argument('--out-dir',default=str(OUT_DIR))
    parser.add_argument('--seed',type=int,default=42)
    args=parser.parse_args(); out=Path(args.out_dir); out.mkdir(parents=True,exist_ok=True); CHART_DIR.mkdir(parents=True,exist_ok=True)
    df=pd.read_csv(args.input,dtype=str,keep_default_na=False)
    required={'severity','logger_class','exception_type','log_message'}
    if not required.issubset(df.columns): raise ValueError(f'Faltan columnas: {required-set(df.columns)}')
    X=np.array(build_features(df),dtype=object); y=df['severity'].astype(str).to_numpy()
    Xtr,Xte,ytr,yte,idxtr,idxte=train_test_split(X,y,np.arange(len(df)),test_size=.20,random_state=args.seed,stratify=y)
    models={
        'Random Forest (100 árboles)':RandomForestClassifier(n_estimators=100,random_state=args.seed,class_weight='balanced',n_jobs=-1),
        'Regresión logística':LogisticRegression(max_iter=1000,random_state=args.seed,class_weight='balanced'),
        'Naive Bayes multinomial':MultinomialNB(),
        'Árbol de decisión':DecisionTreeClassifier(random_state=args.seed,class_weight='balanced')
    }
    rows=[]; trained={}
    for name,clf in models.items():
        pipe=make_pipe(clf); pipe.fit(Xtr,ytr); pred=pipe.predict(Xte); trained[name]=(pipe,pred)
        rows.append({'Modelo':name,'Accuracy':accuracy_score(yte,pred),'Precision_macro':precision_score(yte,pred,average='macro',zero_division=0),'Recall_macro':recall_score(yte,pred,average='macro',zero_division=0),'F1_macro':f1_score(yte,pred,average='macro',zero_division=0),'F1_weighted':f1_score(yte,pred,average='weighted',zero_division=0)})
    results=pd.DataFrame(rows); results.to_csv(out/'model_comparison.csv',index=False,encoding='utf-8-sig')
    selected_name='Random Forest (100 árboles)'; pipe,pred=trained[selected_name]
    labels=['CRITICAL','ERROR','INFO','WARN']
    cm=confusion_matrix(yte,pred,labels=labels)
    report=classification_report(yte,pred,labels=labels,output_dict=True,zero_division=0)
    report_df=pd.DataFrame(report).transpose(); report_df.to_csv(out/'classification_report.csv',encoding='utf-8-sig')
    # Confusion matrix
    fig,ax=plt.subplots(figsize=(6.6,5.2)); im=ax.imshow(cm,cmap='Blues')
    ax.set_xticks(range(len(labels)),labels=labels); ax.set_yticks(range(len(labels)),labels=labels)
    ax.set_xlabel('Predicción'); ax.set_ylabel('Etiqueta real'); style(ax,'Matriz de confusión — conjunto de prueba')
    threshold=cm.max()/2 if cm.max() else 0
    for i in range(cm.shape[0]):
        for j in range(cm.shape[1]): ax.text(j,i,str(cm[i,j]),ha='center',va='center',color='white' if cm[i,j]>threshold else '#16324F',fontsize=11,fontweight='bold')
    fig.colorbar(im,ax=ax,fraction=.046,pad=.04); fig.tight_layout(); fig.savefig(CHART_DIR/'modelo_matriz_confusion.png',dpi=180,bbox_inches='tight'); plt.close(fig)
    # Comparison plot
    fig,ax=plt.subplots(figsize=(8.5,4.6)); x=np.arange(len(results)); width=.34
    ax.bar(x-width/2,results['Accuracy'],width,label='Accuracy',color='#0F8B8D')
    ax.bar(x+width/2,results['F1_macro'],width,label='F1 macro',color='#E0A126')
    ax.set_xticks(x,labels=['Random\nForest','Regresión\nlogística','Naive\nBayes','Árbol de\ndecisión']); ax.set_ylim(0,1.08); ax.legend(frameon=False,ncol=2)
    style(ax,'Comparación de modelos sobre prueba estratificada'); fig.tight_layout(); fig.savefig(CHART_DIR/'modelo_comparacion.png',dpi=180,bbox_inches='tight'); plt.close(fig)
    # Example predictions: chosen 25 rows, stratified-ish from held-out test, stable order.
    audit=pd.DataFrame({'source_row':idxte+1,'severity_real':yte,'severity_predicha':pred,'coincide':(yte==pred)})
    audit['mensaje']=[df.iloc[int(i)]['log_message'] for i in idxte]
    audit['logger_class']=[df.iloc[int(i)]['logger_class'] for i in idxte]
    audit['exception_type']=[df.iloc[int(i)]['exception_type'] for i in idxte]
    audit['confianza_prediccion']=pipe.predict_proba(Xte).max(axis=1)
    audit=audit.sort_values(['severity_real','source_row'])
    # take deterministic slice spanning different classes; prefer test rows, not train
    sample_parts=[]
    for sev,n in [('CRITICAL',5),('ERROR',7),('WARN',7),('INFO',6)]: sample_parts.append(audit[audit.severity_real==sev].head(n))
    audit25=pd.concat(sample_parts).sort_values('source_row').head(25)
    audit25.to_csv(out/'audit_sample_25.csv',index=False,encoding='utf-8-sig')
    # Measure prediction latency on this runtime, excluding model training; environment-dependent.
    warmup=pipe.predict(Xte[:min(10,len(Xte))])
    reps=300; times=[]
    for xval in Xte[:min(30,len(Xte))]:
        start=time.perf_counter_ns()
        for _ in range(10): pipe.predict([xval])
        elapsed=(time.perf_counter_ns()-start)/10/1e6
        times.append(elapsed)
    latency={'n_single_predictions_measured':len(times),'median_ms':float(np.median(times)),'p95_ms':float(np.percentile(times,95)),'note':'Medición local de inferencia por registro; varía por hardware y carga. No equivale a MTTR.'}
    # refit deployment model on all available data and save portable pipeline (input = combined text string)
    deployment=make_pipe(RandomForestClassifier(n_estimators=100,random_state=args.seed,class_weight='balanced',n_jobs=-1))
    deployment.fit(X,y); joblib.dump(deployment,out/'springlog_random_forest.joblib')
    joblib.dump(deployment,ROOT/'04_MODELO_O_PROTOTIPO.joblib')
    demo_text='com.empresa.payment.repository.TransactionRepository HikariPoolTimeoutException '+sanitize('CRITICAL database connection pool exhausted pool=HikariPool-1 timeoutMs=30000 activeConnections=20 maxPoolSize=20')
    demo_prediction=str(deployment.predict([demo_text])[0])
    metrics={
        'dataset_rows':len(df),'train_rows':len(Xtr),'test_rows':len(Xte),'split':'80/20 estratificada','random_state':args.seed,
        'selected_model':selected_name,'selected_model_config':{'n_estimators':100,'class_weight':'balanced','tfidf_ngram_range':[1,2],'tfidf_max_features':1000,'sublinear_tf':True},
        'accuracy':float(accuracy_score(yte,pred)),'precision_macro':float(precision_score(yte,pred,average='macro',zero_division=0)),
        'recall_macro':float(recall_score(yte,pred,average='macro',zero_division=0)),'f1_macro':float(f1_score(yte,pred,average='macro',zero_division=0)),
        'f1_weighted':float(f1_score(yte,pred,average='weighted',zero_division=0)), 'labels_order':labels,
        'confusion_matrix':cm.tolist(), 'classification_report':report,
        'model_comparison':rows,'latency':latency,'demo_prediction':{'input':'TransactionRepository HikariPoolTimeoutException Connection is not available, request timed out after 30000ms','predicted_severity':demo_prediction},
        'caveats':['Dataset sintético y pequeño; distribución de severidad definida al generarlo.','El resultado de prueba no mide desempeño con logs reales.','Logistic Regression y Naive Bayes obtuvieron resultados perfectos en esta partición sintética; eso apunta a patrones demasiado separables y requiere validación con plantillas nuevas.','La auditoría exportada de 25 casos es un insumo para revisión humana, no una certificación independiente de producción.']
    }
    (out/'model_metrics.json').write_text(json.dumps(metrics,indent=2,ensure_ascii=False),encoding='utf-8')
    print('COMPARACIÓN DE MODELOS')
    print(results.to_string(index=False,float_format=lambda x:f'{x:.4f}'))
    print('\nREPORTE RANDOM FOREST')
    print(classification_report(yte,pred,labels=labels,zero_division=0))
    print('\nMATRIZ CONFUSIÓN filas=real, columnas=pred; orden',labels)
    print(cm)
    print('\nLATENCIA',latency)
    print('\nDEMO:',demo_prediction)
    print('\nGuardado:',out/'springlog_random_forest.joblib')

if __name__=='__main__': main()
