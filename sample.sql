function GET_TARIFF_FOR_CONTR(
  ContractID      in dtype. RecordID %Type ,
  TariffRoleA      in   dtype. Name   %Type,
  TariffCodeExt   in   dtype. Name   %Type)
    RETURN TariffDataTable
    PIPELINED
  IS
    TariffData_Record    TariffDataRecord;
    DocRec                doc %RowType;
    ServRec               service_approved %RowType;
    TariffRulesA          dtype. STRING        %Type;
    TariffDataA           tariff_data %RowType;
    rec_id                dtype. RecordID %Type;
    TariffCodeA           dtype. Name   %Type;
    OfficerID             dtype. RecordID %Type;
    CURSOR tariff_data_curs (id_ dtype. RecordID %Type)
      IS
         /*SELECT *
           FROM OPT_DISBUR_PLAN*/
           select * from tariff_data odp
          WHERE (id = id_);
  BEGIN
    select max(id)
      into OfficerID
      from officer
     where amnd_state = stnd.Active
       and UPPER(user_id) = SYS_CONTEXT('USERENV', 'SESSION_USER');

    stnd.START_SESSION(OfficerID,
                       SYS_CONTEXT('USERENV', 'HOST'),
                       SYS_CONTEXT('USERENV', 'MODULE'),
                       null);

    YGDOC(null, DocRec);
    YGSERVICE_APPROVED(null, ServRec);
    TariffRulesA := NULL;
    select tariff_type into  TariffCodeA from tariff
          where amnd_state ='A'
          and tariff_type_ext =  TariffCodeExt
          fetch first 1 row only;
    trf.GET_TARIFF_DATA(ContractID    => ContractID,
                        CDoc          => DocRec,
                        ContraID      => null,
                        TariffRole    => TariffRoleA,
                        TariffCode    => TariffCodeA,
                        SearchRules   => null,
                        TariffRules   => TariffRulesA,
                        TariffData    =>  TariffDataA);

     IF TariffDataA.id IS NOT NULL then
        OPEN tariff_data_curs (TariffDataA.id);
         LOOP
            FETCH tariff_data_curs INTO TariffData_Record;

            EXIT WHEN tariff_data_curs%NOTFOUND;
            PIPE ROW (TariffData_Record);
         END LOOP;
     END IF;
     RETURN;
     stnd.finish_session;
  END;
