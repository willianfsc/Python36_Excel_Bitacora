#!/usr/bin/env python3.6


###################################
#     IMPORTACAO  DE LIBRARIES
###################################
import xlsxwriter
import datetime
import psycopg2
import time

#####  VARIAVEIS DE CONTROLE  #####
limparq=0

#####  VERIFICA O RANGE DO MES ATUAL  #####
primeiro = datetime.datetime.today().replace(day=1, hour=00, minute=00, second=00).strftime('%Y-%m-%d %H:%M:%S')
hoje     = datetime.datetime.today().strftime('%Y-%m-%d %H:%M:%S')
monthout = datetime.datetime.today().strftime('%y%m%d')

print ("PRIMEIRO: {}".format(primeiro))
print ("HOJE    : {}".format(hoje))


##### VERIFICA A LISTA DE JOBS CADASTRADOS #####
with open('cip_jobs_int.txt') as jobs:
    for job in jobs:
        job = job.replace("\n", "")
#        print (job)

        #####  CRIA A QUERY  #####
        query="SELECT job_mem_name, start_time, end_time, end_time - start_time as elapsed_time FROM runinfo_history WHERE job_mem_name = '"+job+"' AND start_time >= '"+primeiro+"'  AND start_time <= '"+hoje+"' GROUP BY job_mem_name, start_time, end_time ORDER BY start_time;"

#        print (query)

        ####  CONEXAO E INTERACAO COM A BASE DE DADOS POSTGRESQL  ####
        try:
            conn = psycopg2.connect("dbname='em900' user='emuser' host='172.28.12.66' password='manager'")
        except:
            print ("I am unable to connect to the database")

        cur = conn.cursor()
        cur.execute(query)

        #### SAIDA DA QUERY ESTA ARMAZENADA NA VARIAVEL ROWS  ####
        rows = cur.fetchall()
        for row in rows:
            delay = row[3]
            if (delay.days > 0):
                out = str(delay).replace(" days, ", ":")
            else:
                out = "0" + str(delay)

            #outputquery="{},{},{},{}".format(row[0],row[1],row[2],row[3])
            outputquery="{},{},{},{}".format(row[0],row[1],row[2],out)
            print (outputquery)

            #### SALVAR A LINHA EM ARQUIVO    #####
            if ( limparq == 0):
                f = open( 'ctm_bitacora_'+monthout+'.csv', 'w' )         ##### SALVA APENAS UMA LINHA NO ARQUIVO OU LIMPA O ARQUIVO ANTES DE GRAVAR
                limparq=1
            f = open( 'ctm_bitacora_'+monthout+'.csv', 'a' )             ##### EFEUTA O APPEND DA LINHA NO ARQUIVO
            f.write( outputquery + '\n' )
            f.close()

#### FECHAR AS COMUNICACOES COM O POSTGRESQL  ####
cur.close()
conn.close()
