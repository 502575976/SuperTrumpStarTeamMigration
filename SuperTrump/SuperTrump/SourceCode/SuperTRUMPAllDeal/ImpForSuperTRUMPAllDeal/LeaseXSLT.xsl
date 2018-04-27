<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
xmlns:ms="urn:schemas-microsoft-com:xslt"
xmlns:dt="urn:schemas-microsoft-com:datatypes">
  <xsl:template match="NewDataSet">
    <!-- These will be activated in case of webservice call and for DEsktop ST these Tags will be Commented-->
    <PRM_INFO>
      <PRM_FILE>
        <AD_HOC_QUERY>
          <SuperTRUMP>
            <Transaction query="QUI_001" id="TRAN1">
              <Initialize/>
              <ReadFile>
                <xsl:attribute name="filename">
                  <xsl:value-of select="TemplateMapping/TEMPLATENAME"/>
                </xsl:attribute>
              </ReadFile>
              <TransactionStartDate>
                <xsl:if test="AccountScheduleFeed/TRANSACTION_START_DATE!=''">
                  <xsl:value-of select="substring(substring-after(substring-after(AccountScheduleFeed/TRANSACTION_START_DATE,'/'),'/'),1,4)" />-<xsl:value-of select="substring-before(AccountScheduleFeed/TRANSACTION_START_DATE,'/')" />-<xsl:value-of select="substring-before(substring-after(AccountScheduleFeed/TRANSACTION_START_DATE,'/'),'/')" />
                </xsl:if>
              </TransactionStartDate>
              <TermInMonths>
                <xsl:value-of select="AccountScheduleFeed/TERM"/>
              </TermInMonths>
              <xsl:for-each select="AccountScheduleFeed">
                <xsl:if test="FREQUENCY='M'">
                  <Periodicity>Monthly</Periodicity>
                </xsl:if>
                <xsl:if test="FREQUENCY='A'">
                  <Periodicity>Annually</Periodicity>
                </xsl:if>
                <xsl:if test="FREQUENCY='Q'">
                  <Periodicity>Quaterly</Periodicity>
                </xsl:if>
              </xsl:for-each>
              <xsl:for-each select="AccountScheduleFeed">
                <xsl:if test="ADV_ARR='ARR'">
                  <PaymentTiming>Arrears</PaymentTiming>
                </xsl:if>
                <xsl:if test="ADV_ARR='ADV'">
                  <PaymentTiming>Advance</PaymentTiming>
                </xsl:if>
              </xsl:for-each>
              <xsl:choose>
                <xsl:when test="AccountScheduleFeed/PRODUCT='SGLINV'">
                  <AutoRVI>true</AutoRVI>
                  <RVIRate>
                    <xsl:value-of select="format-number(AccountScheduleFeed/RESIDUAL_INSURANCE_PREMIUM_PCT,'#.##') div 100"/>
                  </RVIRate>
                </xsl:when>
                <xsl:otherwise>
                  <AutoRVI>false</AutoRVI>
                  <RVIRate/>
                </xsl:otherwise>
              </xsl:choose>
              <CommencementDate>
                <xsl:if test="AccountScheduleFeed/COMMENCEMENT_DATE!=''">
                  <xsl:value-of select="substring(substring-after(substring-after(AccountScheduleFeed/COMMENCEMENT_DATE,'/'),'/'),1,4)" />-<xsl:value-of select="substring-before(AccountScheduleFeed/COMMENCEMENT_DATE,'/')" />-<xsl:value-of select="substring-before(substring-after(AccountScheduleFeed/COMMENCEMENT_DATE,'/'),'/')" />
                </xsl:if>
              </CommencementDate>
              <Features>
                <!--<AllowSubsidies>true</AllowSubsidies>
						        <AllowAccounting>true</AllowAccounting>
						        <AllowFederalTaxes>true</AllowFederalTaxes>
					          <AllowPeriodicIncomeExpense>true</AllowPeriodicIncomeExpense>
                    <AllowStateTaxes>true</AllowStateTaxes>
						        <AllowAssetAssociation>true</AllowAssetAssociation>-->
                <AllowSecurityDeposits>true</AllowSecurityDeposits>
              </Features>
              <LeaseID>
                <xsl:value-of select="AccountScheduleFeed/ACCOUNT_SCHEDULE_NBR"/>
                <xsl:value-of select="AccountScheduleFeed/IDMS_REGION"/>
              </LeaseID>
              <xsl:choose>
                <xsl:when test="AccountScheduleFeed/PRODUCT='MEREG' or AccountScheduleFeed/PRODUCT='MEQMUN' or AccountScheduleFeed/PRODUCT='MEOQSI' or AccountScheduleFeed/PRODUCT='MEGMUN'">
                  <LendingLoans>
                    <LendingLoan index="{position()-1}">
                      <CashflowSteps>
                        <CashflowStep index="{position()-1}">
                          <Type>Funding</Type>
                          <Amount>
                            -<xsl:value-of select="format-number(sum(//NewDataSet/AssetLevelFeed/OEC_ON_ASSET),'#.##')"/>
                          </Amount>
                        </CashflowStep>
                        <xsl:choose>
                          <xsl:when test="number(substring(substring-after(substring-after(//NewDataSet/AccountScheduleFeed/COMMENCEMENT_DATE,'/'),'/'),1,4))!=number(substring(substring-after(substring-after(//NewDataSet/AccountScheduleFeed/TRANSACTION_START_DATE,'/'),'/'),1,4)) or number(substring-before(//NewDataSet/AccountScheduleFeed/COMMENCEMENT_DATE,'/'))!=number(substring-before(//NewDataSet/AccountScheduleFeed/TRANSACTION_START_DATE,'/')) or number(substring-before(substring-after(//NewDataSet/AccountScheduleFeed/COMMENCEMENT_DATE,'/'),'/'))!=number(substring-before(substring-after(//NewDataSet/AccountScheduleFeed/TRANSACTION_START_DATE,'/'),'/'))">
                            <CashflowStep index="{position()}">
                              <NumberOfPayments>1</NumberOfPayments>
                              <Periodicity>Stub</Periodicity>
                              <Type>
                                <xsl:choose>
                                  <xsl:when test="//NewDataSet/AccountScheduleFeed/INTERIM_RENT =0 or //NewDataSet/AccountScheduleFeed/INTERIM_RENT ='' or //NewDataSet/AccountScheduleFeed/INTERIM_RENT =NULL">
                                    Payment
                                  </xsl:when>
                                  <xsl:otherwise>
                                    Stub
                                  </xsl:otherwise>
                                </xsl:choose>
                              </Type>
                              <IsStub>true</IsStub>
                              <Amount>
                                <xsl:choose>
                                  <xsl:when test="//NewDataSet/AccountScheduleFeed/INTERIM_RENT !=0 or //NewDataSet/AccountScheduleFeed/INTERIM_RENT !='' or //NewDataSet/AccountScheduleFeed/INTERIM_RENT !=NULL">
                                    <xsl:value-of select="//NewDataSet/AccountScheduleFeed/INTERIM_RENT"/>
                                  </xsl:when>
                                  <xsl:otherwise>
                                    0
                                  </xsl:otherwise>
                                </xsl:choose>
                              </Amount>
                            </CashflowStep>
                            <xsl:for-each select="//NewDataSet/StreamFeed">
                              <CashflowStep index="{position()+1}">
                                <NumberOfPayments>
                                  <xsl:value-of select="NO_OF_PAYMENTS"/>
                                </NumberOfPayments>
                                <Type>Payment</Type>
                                <Periodicity>Monthly</Periodicity>
                                <Amount>
                                  <xsl:value-of select="AMOUNT"/>
                                </Amount>
                              </CashflowStep>
                            </xsl:for-each>
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:for-each select="//NewDataSet/StreamFeed">
                              <CashflowStep index="{position()}">
                                <NumberOfPayments>
                                  <xsl:value-of select="NO_OF_PAYMENTS"/>
                                </NumberOfPayments>
                                <Type>Payment</Type>
                                <Periodicity>Monthly</Periodicity>
                                <Amount>
                                  <xsl:value-of select="AMOUNT"/>
                                </Amount>
                              </CashflowStep>
                            </xsl:for-each>
                          </xsl:otherwise>
                        </xsl:choose>
                      </CashflowSteps>
                    </LendingLoan>
                  </LendingLoans>
                </xsl:when>
                <xsl:otherwise>
                  <Assets>
                    <xsl:for-each select="AssetLevelFeed">
                      <Asset index="{position()-1}">
                        <Cost>
                          <xsl:value-of select="OEC_ON_ASSET"/>
                        </Cost>
                        <DeliveryDate>
                          <xsl:if test="DATE_IN_SERVICE!=''">
                            <xsl:value-of select="substring(substring-after(substring-after(DATE_IN_SERVICE,'/'),'/'),1,4)" />-<xsl:value-of select="substring-before(DATE_IN_SERVICE,'/')" />-<xsl:value-of select="substring-before(substring-after(DATE_IN_SERVICE,'/'),'/')" />
                          </xsl:if>
                        </DeliveryDate>
                        <FundingDate>
                          <xsl:if test="FUNDING_DATE!=''">
                            <xsl:value-of select="substring(substring-after(substring-after(FUNDING_DATE,'/'),'/'),1,4)" />-<xsl:value-of select="substring-before(FUNDING_DATE,'/')" />-<xsl:value-of select="substring-before(substring-after(FUNDING_DATE,'/'),'/')" />
                          </xsl:if>
                        </FundingDate>
                        <SalesTax>
                          <xsl:value-of select="CAPITALIZED_FEES"/>
                        </SalesTax>
                        <ResidualKeptAsAPercent>false</ResidualKeptAsAPercent>
                        <Residual>
                          <xsl:value-of select="RESEDUAL_AMOUNT"/>
                        </Residual>
                        <GuaranteeType>Lessee</GuaranteeType>
                        <xsl:if test="position()-1 = 0">
                          <GuaranteedAmount>
                            <xsl:value-of select="//NewDataSet/AccountScheduleFeed/GUARANTEED_RESIDUAL"/>
                          </GuaranteedAmount>
                        </xsl:if>
                        <GuaranteeThirdPartyIsPut>false</GuaranteeThirdPartyIsPut>
                        <FederalDepreciation>
                          <Method>
                            <xsl:value-of select="//NewDataSet/Depriciation/METHOD"/>
                          </Method>
                          <!--<PercentDepreciable />-->
                          <Term>
                            <xsl:value-of select="DEPRECIATION_CNT div 12"/>
                          </Term>
                          <Salvage>
                            <xsl:value-of select="SALVAGE_AMT"/>
                          </Salvage>
                        </FederalDepreciation>
                        <StateDepreciation>
                          <Method>
                            <xsl:value-of select="//NewDataSet/Depriciation/METHOD"/>
                          </Method>
                          <Term>
                            <xsl:value-of select="DEPRECIATION_CNT div 12"/>
                          </Term>
                          <Salvage>
                            <xsl:value-of select="SALVAGE_AMT"/>
                          </Salvage>
                        </StateDepreciation>
                      </Asset>
                    </xsl:for-each>
                  </Assets>
                </xsl:otherwise>
              </xsl:choose>
              <Fees>
                <Fee index="{position()-1}">
                  <Description>Income Fee</Description>
                  <IsAnExpense>false</IsAnExpense>
                  <KeptAsAPercent>false</KeptAsAPercent>
                  <Amount>
                    <!--<xsl:value-of select="AccountScheduleFeed/INCOME_TYPE_FEES"/>-->
                    <xsl:choose>
                      <xsl:when test="contains(AccountScheduleFeed/INCOME_TYPE_FEES,'-')">
                        <xsl:value-of select="AccountScheduleFeed/INCOME_TYPE_FEES * -1"/>
                      </xsl:when>
                      <xsl:otherwise>
                        <xsl:value-of select="AccountScheduleFeed/INCOME_TYPE_FEES"/>
                      </xsl:otherwise>
                    </xsl:choose>
                  </Amount>
                  <IsLesseeObligation>false</IsLesseeObligation>
                </Fee>
                <Fee index="{position()}">
                  <Description>Expense Fee</Description>
                  <IsAnExpense>true</IsAnExpense>
                  <KeptAsAPercent>false</KeptAsAPercent>
                  <Amount>
                    <xsl:value-of select="AccountScheduleFeed/EXPENSE_TYPE_FEES"/>
                  </Amount>
                  <IsLesseeObligation>false</IsLesseeObligation>
                </Fee>
              </Fees>
              <SecurityDeposits>
                <Method>Specified</Method>
                <SpecifiedSecurityDeposits>
                  <xsl:value-of select="AccountScheduleFeed/SECURITY_DEPOSIT_AMOUNT"/>
                </SpecifiedSecurityDeposits>
              </SecurityDeposits>              
              
              <xsl:choose>
                <xsl:when test="AccountScheduleFeed/PRODUCT!='MEREG' and AccountScheduleFeed/PRODUCT!='MEQMUN' and AccountScheduleFeed/PRODUCT!='MEOQSI' and AccountScheduleFeed/PRODUCT!='MEGMUN'">
                  <Rents>
                    <Rent index="{position()-1}">
                      <AssociationIndex>0</AssociationIndex>
                      <AddResidualAsFunding>false</AddResidualAsFunding>
                      <CashflowSteps>
                        <xsl:choose>
                          <xsl:when test="number(substring(substring-after(substring-after(//NewDataSet/AccountScheduleFeed/COMMENCEMENT_DATE,'/'),'/'),1,4))!=number(substring(substring-after(substring-after(//NewDataSet/AccountScheduleFeed/TRANSACTION_START_DATE,'/'),'/'),1,4)) or number(substring-before(//NewDataSet/AccountScheduleFeed/COMMENCEMENT_DATE,'/'))!=number(substring-before(//NewDataSet/AccountScheduleFeed/TRANSACTION_START_DATE,'/')) or number(substring-before(substring-after(//NewDataSet/AccountScheduleFeed/COMMENCEMENT_DATE,'/'),'/'))!=number(substring-before(substring-after(//NewDataSet/AccountScheduleFeed/TRANSACTION_START_DATE,'/'),'/'))">
                            <CashflowStep index="{position()-1}">
                              <NumberOfPayments>1</NumberOfPayments>
                              <Periodicity>Stub</Periodicity>
                              <Type>
                                <xsl:choose>
                                  <xsl:when test="//NewDataSet/AccountScheduleFeed/INTERIM_RENT =0 or //NewDataSet/AccountScheduleFeed/INTERIM_RENT ='' or //NewDataSet/AccountScheduleFeed/INTERIM_RENT =NULL">
                                    Payment
                                  </xsl:when>
                                  <xsl:otherwise>Stub</xsl:otherwise>
                                </xsl:choose>
                              </Type>
                              <IsStub>true</IsStub>
                              <Amount>
                                <xsl:choose>
                                  <xsl:when test="//NewDataSet/AccountScheduleFeed/INTERIM_RENT !=0 or //NewDataSet/AccountScheduleFeed/INTERIM_RENT !='' or //NewDataSet/AccountScheduleFeed/INTERIM_RENT !=NULL">
                                    <xsl:value-of select="//NewDataSet/AccountScheduleFeed/INTERIM_RENT"/>
                                  </xsl:when>
                                  <xsl:otherwise>0</xsl:otherwise>
                                </xsl:choose>
                              </Amount>
                            </CashflowStep>
                            <xsl:choose>
                              <xsl:when test="AccountScheduleFeed/TERM != sum(//NewDataSet/StreamFeed/NO_OF_PAYMENTS)">
                                <CashflowStep index="{position()}">
                                  <NumberOfPayments>1</NumberOfPayments>
                                  <Type>Payment</Type>
                                  <Periodicity>Monthly</Periodicity>
                                  <Amount>0</Amount>
                                </CashflowStep>
                                <xsl:for-each select="StreamFeed">
                                  <CashflowStep index="{position()+1}">
                                    <NumberOfPayments>
                                      <xsl:value-of select="NO_OF_PAYMENTS"/>
                                    </NumberOfPayments>
                                    <Type>Payment</Type>
                                    <Periodicity>Monthly</Periodicity>
                                    <Amount>
                                      <xsl:value-of select="AMOUNT"/>
                                    </Amount>
                                  </CashflowStep>
                                </xsl:for-each>
                              </xsl:when>
                              <xsl:otherwise>
                                <xsl:for-each select="StreamFeed">
                                  <CashflowStep index="{position()}">
                                    <NumberOfPayments>
                                      <xsl:value-of select="NO_OF_PAYMENTS"/>
                                    </NumberOfPayments>
                                    <Type>Payment</Type>
                                    <Periodicity>Monthly</Periodicity>
                                    <Amount>
                                      <xsl:value-of select="AMOUNT"/>
                                    </Amount>
                                  </CashflowStep>
                                </xsl:for-each>
                              </xsl:otherwise>
                            </xsl:choose>
                          </xsl:when>
                          <xsl:otherwise>
                            <xsl:choose>
                              <xsl:when test="AccountScheduleFeed/TERM != sum(//NewDataSet/StreamFeed/NO_OF_PAYMENTS)">
                                <CashflowStep index="{position()-1}">
                                  <NumberOfPayments>1</NumberOfPayments>
                                  <Type>Payment</Type>
                                  <Periodicity>Monthly</Periodicity>
                                  <Amount>0</Amount>
                                </CashflowStep>
                                <xsl:for-each select="StreamFeed">
                                  <CashflowStep index="{position()}">
                                    <NumberOfPayments>
                                      <xsl:value-of select="NO_OF_PAYMENTS"/>
                                    </NumberOfPayments>
                                    <Type>Payment</Type>
                                    <Periodicity>Monthly</Periodicity>
                                    <Amount>
                                      <xsl:value-of select="AMOUNT"/>
                                    </Amount>
                                  </CashflowStep>
                                </xsl:for-each>
                              </xsl:when>
                              <xsl:otherwise>
                                <xsl:for-each select="StreamFeed">
                                  <CashflowStep index="{position()-1}">
                                    <NumberOfPayments>
                                      <xsl:value-of select="NO_OF_PAYMENTS"/>
                                    </NumberOfPayments>
                                    <Type>Payment</Type>
                                    <Periodicity>Monthly</Periodicity>
                                    <Amount>
                                      <xsl:value-of select="AMOUNT"/>
                                    </Amount>
                                  </CashflowStep>
                                </xsl:for-each>
                              </xsl:otherwise>
                            </xsl:choose>
                          </xsl:otherwise>
                        </xsl:choose>
                      </CashflowSteps>
                    </Rent>
                  </Rents>
                </xsl:when>
              </xsl:choose>
              <LessorLender>
                <UseAMT>false</UseAMT>
                <doStateNOL>false</doStateNOL>
              </LessorLender>
              <GEdata>
                <GEBusiness>VF</GEBusiness>
                <GEProduct>
                  <xsl:value-of select="ProductMapping/ST_PRODUCT_NAME"/>
                </GEProduct>
                <NoTaxBase>false</NoTaxBase>
                <ResidualUpside>
                  <xsl:value-of select="//NewDataSet/RESIDUALMAPPING/CAP_MARKET_ADDER"/>
                </ResidualUpside>
                <MoneyCostDate>
                  <xsl:value-of select="substring(substring-after(substring-after(AccountScheduleFeed/BOOKING_DATE,'/'),'/'),1,4)" />-<xsl:value-of select="substring-before(AccountScheduleFeed/BOOKING_DATE,'/')" />-<xsl:value-of select="substring-before(substring-after(AccountScheduleFeed/BOOKING_DATE,'/'),'/')" />
                </MoneyCostDate>
              </GEdata>
              <IsTemplate>false</IsTemplate>
              <Writefile>
                <xsl:attribute name="filename">\\comfinciohnas.comfin.ge.com\CIOHShared$\PRICING_ROI\DEV\ST_ALL_DEAL\<xsl:value-of select="AccountScheduleFeed/ACCOUNT_SCHEDULE_NBR"/>.prm</xsl:attribute>
              </Writefile>
            </Transaction>
          </SuperTRUMP>
        </AD_HOC_QUERY>
      </PRM_FILE>
    </PRM_INFO>
  </xsl:template>
</xsl:stylesheet>
