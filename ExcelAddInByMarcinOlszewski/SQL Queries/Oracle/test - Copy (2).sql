select 
s.primary_transaction_id, s.source_feeder_code,s.ship_order_flag, s.FISCAL_MONTH,s.GTM_ROUTE,s.PRODUCT_LINE,s.SALES_ORG_COUNTRY, fl.FLIP_COUNTRY, s.SALES_ORG_REGION,s.SALES_ORG_SUB_REGION,
s.SITE_C_COUNTRY,s.SOURCE_REGION,s.CUSTOMER_SEGMENT,s.CUSTOMER_TARGET_SEGMENT,
s.LEAD_SLS_COV_SEG,s.OFFER_SUB_TYPE,s.OFFERING_TYPE,s.TRANSACTION_TYPE,
s.SITE_A_Country,s.SITE_B_Country,s.DEAL_TYPE,s.PARTNER_TIER,s.AREA_CODE,s.CREDITING_TYPE ,s.CHANNEL_SALES_MOTION,
s.SALES_CHANNEL_CODE ,s.TABLE_SOURCE,
sum(s.AMOUNT_USD)
from  (
select primary_transaction_id, source_feeder_code,ship_order_flag, FISCAL_MONTH,GTM_ROUTE,PRODUCT_LINE,SALES_ORG_COUNTRY, SALES_ORG_REGION,SALES_ORG_SUB_REGION,
SITE_C_COUNTRY,SOURCE_REGION,CUSTOMER_SEGMENT,CUSTOMER_TARGET_SEGMENT,
LEAD_SLS_COV_SEG,OFFER_SUB_TYPE,OFFERING_TYPE,TRANSACTION_TYPE,
SITE_A_Country,SITE_B_Country,DEAL_TYPE,PARTNER_TIER,AREA_CODE,CREDITING_TYPE ,CHANNEL_SALES_MOTION,
SALES_CHANNEL_CODE, site_c_account_level_id,CFEED_RELEVANT,'OUTB TABLE' as TABLE_SOURCE, amount_usd
from siqp_omega_trans_Sum where fiscal_month >202400
union all
select primary_transaction_id, source_feeder_code,ship_order_flag, FISCAL_MONTH,GTM_ROUTE,PRODUCT_LINE,SALES_ORG_COUNTRY, SALES_ORG_REGION,SALES_ORG_SUB_REGION,
SITE_C_COUNTRY,SOURCE_REGION,CUSTOMER_SEGMENT,CUSTOMER_TARGET_SEGMENT,
LEAD_SLS_COV_SEG,OFFER_SUB_TYPE,OFFERING_TYPE,TRANSACTION_TYPE,
SITE_A_Country,SITE_B_Country,DEAL_TYPE,PARTNER_TIER,AREA_CODE,CREDITING_TYPE ,CHANNEL_SALES_MOTION,
SALES_CHANNEL_CODE,  site_c_account_level_id, '' as CFEED_RELEVANT, 'BLOCK TABLE' as TABLE_SOURCE,amount_usd 
from tbl_cfeed_blocked_data  where fiscal_month >202400  )  s
inner join tbl_global_ref ref on (  s.PRODUCT_LINE = ref.prod_line_code and ref.cfeed_eligible = 'Y'  and ref.fiscal_year =  substr(fiscal_month,1,4)  and fiscal_half = case  when  substr(fiscal_month,5,2)  >6 then '2H' Else '1H' end   )
left join TBL_CTRY_BY_ACCOUNT fl on ( fl.fiscal_year =  substr(fiscal_month,1,4) and s.site_c_account_level_id = fl.ACCOUNT_ID and s.sales_org_country = fl.TRANSACTION_COUNTRY)
where crediting_type ='N'   and  (   ( s.ship_order_flag='O')  or
(s.ship_order_flag ='P' and source_feeder_code IN ('MI', 'ME','SO') ) or
( s.ship_order_flag ='G'  and s.TRANSACTION_TYPE='SELL OUT' AND s.SOURCE_FEEDER_CODE='PSO')
or (  s.SHIP_ORDER_FLAG ='M' AND TRIM(s.CFEED_RELEVANT)='Y')
)
group by 
s.primary_transaction_id, s.source_feeder_code,s.ship_order_flag, s.FISCAL_MONTH,s.GTM_ROUTE,s.PRODUCT_LINE,s.SALES_ORG_COUNTRY, fl.FLIP_COUNTRY, s.SALES_ORG_REGION,s.SALES_ORG_SUB_REGION,
s.SITE_C_COUNTRY,s.SOURCE_REGION,s.CUSTOMER_SEGMENT,s.CUSTOMER_TARGET_SEGMENT,
s.LEAD_SLS_COV_SEG,s.OFFER_SUB_TYPE,s.OFFERING_TYPE,s.TRANSACTION_TYPE,
s.SITE_A_Country,s.SITE_B_Country,s.DEAL_TYPE,s.PARTNER_TIER,s.AREA_CODE,s.CREDITING_TYPE ,s.CHANNEL_SALES_MOTION,
s.SALES_CHANNEL_CODE,s.TABLE_SOURCE FETCH FIRST 100 ROWS ONLY