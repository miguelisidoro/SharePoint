import * as strings from 'InOfficeSpFxWebPartStrings';

import {
  HttpClient,
  HttpClientConfiguration,
  HttpClientResponse,
  IHttpClientOptions,
  ISPHttpClientOptions,
  SPHttpClient,
  SPHttpClientResponse
} from "@microsoft/sp-http";
 import {
//   ItemAddResult,
//   ItemUpdateResult,
//   PagedItemCollection,
//   Web,
   sp
} from "@pnp/sp";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { PagedItemCollection } from "@pnp/sp/items";
import {IInOfficeAppointment} from "../interfaces/IInOfficeAppointment"

// import { format, parseISO } from "date-fns/esm";

import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
// import { ICanal } from "./ICanal";
// import { ICentro } from "./ICentro";
// import { ICentroCusto } from "./ICentroCusto";
// import { IClassificacaoContabil } from "./IClassificacaoContabil";
// import { IConfigurationValue } from "./IConfigurationValue";
// import { IContaGastos } from "./IContaGastos";
// import { IEmpresa } from "./IEmpresa";
// import { IFornecedor } from "./IFornecedor";
// import { IGrupoCompradores } from "./IGrupoCompradores";
// import { IImobilizado } from "./IImobilizado";
// import { IListaPrecos } from "./IListaPrecos";
// import { IMaterial } from "./IMaterial";
// import { IMoeda } from "./IMoeda";
// import { IOrdem } from "./IOrdem";
// import { IPais } from "./IPais";
// import { IPedidoCompra } from "./IPedidoCompra";
// import { IPedidoCompraDetalhe } from "./IPedidoCompraDetalhe";
// import { IRegiaoVendas } from "./IRegiaoVendas";
// import { ISector } from "./ISector";
// import { ISegmentoMarca } from "./ISegmentoMarca";
// import { IUnidadeMedida } from "./IUnidadeMedida";
// import { Log } from "@microsoft/sp-core-library";
// import { SPComponentLoader } from "@microsoft/sp-loader";
// import { ServiceScope } from "@microsoft/sp-core-library";
 import { WebPartContext } from "@microsoft/sp-webpart-base";
// import { errorType } from "./IEnumErrorType";
// import { estadoPedido } from "./../spservices/IEnumEstadosPedido";
import { graph } from "@pnp/graph";
// import { pt } from "date-fns/locale";
// import { IGrupoMercadorias } from "./IGrupoMercadorias";
// import { IMarca } from "./IMarca";

// const CONFIGURATION_LIST: string = "Configuration List";
const LOG_SOURCE: string = "In Office Appointment";
const IN_OFFICE_APPOINTMENT_LIST_NAME: string = "InOfficeAppointments";
// const PEDIDO_COMPRA_CONF_KEY = "Pedidocompra";
// const PEDIDO_COMPRA_DETALHE_CONF_KEY = "Pedidocompradetalhe";

// const EMPRESA_CONF_KEY = "Empresa";
// const MARCA_CONF_KEY = "Marca";
// const FORNECEDOR_CONF_KEY = "Fornecedor";
// const GRUPOCOMPRADORES_CONF_KEY = "GrupoCompradores";
// const GRUPOMERCADORIAS_CONF_KEY = "GrupoMercadorias";
// const CENTRO_CONF_KEY = "Centro";
// const MATERIAL_CONF_KEY = "Material";
// const MOEDA_CONF_KEY = "Moeda";
// const UNIDADE_MEDIDA_CONF_KEY = "UnidadeMedida";
// const CLASSIFICACAO_CONTABIL_CONF_KEY = "ClassificacaoContabil";
// const IMOBILIZADO_CONF_KEY = "Imobilizado";
// const CANAL_CONF_KEY = "Canal";
// const LISTAPRECOS_CONF_KEY = "ListaPrecos";
// const SECTOR_CONF_KEY = "Sector";
// const REGIAO_VENDAS_CONF_KEY = "RegiaoVendas";
// const PAIS_CONF_KEY = "Pais";
// const CONTAGASTOS_CONF_KEY = "ContaGastos";
// const ORDEM_CONF_KEY = "Ordem";
// const SEGMENTOMARCA_CONF_KEY = "SegmentoDeMarca";
// const CENTROCUSTOS_CONF_KEY = "CentroCustos";
// const APPROVAL_FLOW_ENDPOINT_KEY = "APPROVAL_FLOW_ENDPOINT";

// const DEFAULT_PERSONA_IMG_HASH: string = "7ad602295f8386b7615b582d87bcc294";
// const DEFAULT_IMAGE_PLACEHOLDER_HASH: string =
//   "4a48f26592f4e1498d7a478a4c48609c";
// const MD5_MODULE_ID: string = "8494e7d7-6b99-47b2-a741-59873e42f16f";
// const PROFILE_IMAGE_URL: string =
//   "/_layouts/15/userphoto.aspx?size=M&accountname=";

// Class Services
export default class spservices {
  constructor(private _context: WebPartContext | ApplicationCustomizerContext) {
    // Setup Context to PnPjs and MSGraph
    sp.setup({
      spfxContext: this._context
    });

    graph.setup({
      spfxContext: this._context
    });

    this.onInit();
  }
  // OnInit Function
  private async onInit() {}

   // /**
  //  * Gets In Office Appointments
  //  * @param sortField
  //  * @param ascending
  //  * @returns pedidos compra
  //  */
   public async getInOfficeAppointments(
    sortField: string,
    ascending: boolean
  ): Promise<PagedItemCollection<IInOfficeAppointment[]>> {
    console.log("Start getInOfficeAppointments..");
    Log.verbose(
      LOG_SOURCE,
      "Start getInOfficeAppointments..",
      this._context.serviceScope
    );
    
    try {
          const results = await sp.web.lists
            .getByTitle(IN_OFFICE_APPOINTMENT_LIST_NAME)
            .items.select(
              "Id",
              "Colaborador",
              "Data",
              "Notas",
              "ContactosPr_x00f3_ximos"
            )
            .orderBy(`${sortField}`, ascending)
            .top(10)
            .getPaged();
          Log.verbose(
            LOG_SOURCE,
            "End getInOfficeAppointments..",
            this._context.serviceScope
          );
          return results;
      }
    } catch (error) {
      console.log(error);
      Log.error(LOG_SOURCE, error, this._context.serviceScope);
      throw new Error(error.message);
    }
  }

 // public getUnidadeMedida = async (): Promise<IUnidadeMedida[]> => {
  //   try {
  //     const unidadeMedidaListId = await this.getConfigurationValue(
  //       UNIDADE_MEDIDA_CONF_KEY
  //     );
  //     if (unidadeMedidaListId) {
  //       const results: IUnidadeMedida[] = await sp.web.lists
  //         .getById(unidadeMedidaListId)
  //         .items.select("UnidadeMedida", "DescricaoItem")
  //         .usingCaching()
  //         .orderBy("DescricaoItem")
  //         .get();
  //       return results && results.length > 0
  //         ? results
  //         : ({} as IUnidadeMedida[]);
  //     } else {
  //       const error: Error = new Error(
  //         `${strings.getUnidadeMedidaErrorMessage}`
  //       );
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error.message);
  //   }
  // };

  // public getUnidadeMedidaById = async (
  //   id: string
  // ): Promise<IUnidadeMedida[]> => {
  //   try {
  //     const unidadeMedidaListId = await this.getConfigurationValue(
  //       UNIDADE_MEDIDA_CONF_KEY
  //     );
  //     if (unidadeMedidaListId) {
  //       const results: IUnidadeMedida[] = await sp.web.lists
  //         .getById(unidadeMedidaListId)
  //         .items.getById(Number(id))
  //         .select("UnidadeMedida", "DescricaoItem")
  //         .usingCaching()
  //         .get();
  //       return results && results.length > 0
  //         ? results
  //         : ({} as IUnidadeMedida[]);
  //     } else {
  //       const error: Error = new Error(
  //         `${strings.getUnidadeMedidaErrorMessage}`
  //       );
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error.message);
  //   }
  // };

  // public getClassificacaoContabil = async (): Promise<
  //   IClassificacaoContabil[]
  // > => {
  //   try {
  //     const classificacaoContabilListId = await this.getConfigurationValue(
  //       CLASSIFICACAO_CONTABIL_CONF_KEY
  //     );
  //     if (classificacaoContabilListId) {
  //       const results: IClassificacaoContabil[] = await sp.web.lists
  //         .getById(classificacaoContabilListId)
  //         .items.select("ClassificacaoContabil", "DescricaoItem")
  //         .usingCaching()
  //         .orderBy("ClassificacaoContabil")
  //         .get();
  //       return results && results.length > 0
  //         ? results
  //         : ({} as IClassificacaoContabil[]);
  //     } else {
  //       const error: Error = new Error(
  //         `${strings.getUnidadeMedidaErrorMessage}`
  //       );
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error.message);
  //   }
  // };

  // public getMoedas = async (): Promise<IMoeda[]> => {
  //   try {
  //     const moedaListId = await this.getConfigurationValue(MOEDA_CONF_KEY);
  //     if (moedaListId) {
  //       const results: IMoeda[] = await sp.web.lists
  //         .getById(moedaListId)
  //         .items.select("Moeda", "DescricaoItem")
  //         .usingCaching()
  //         .orderBy("DescricaoItem")
  //         .get();
  //       return results && results.length > 0 ? results : ({} as IMoeda[]);
  //     } else {
  //       const error: Error = new Error(`${strings.GetMoedasErrorMessage}`);
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error.message);
  //   }
  // };

  // // Get List Configuration Value

  // /**
  //  * Gets configuration value
  //  * @param configurationKey
  //  * @returns configuration value
  //  */
  // public async getConfigurationValue(
  //   configurationKey: string
  // ): Promise<string> {
  //   try {
  //     const configurationValue: IConfigurationValue[] = await sp.web.lists
  //       .getByTitle(CONFIGURATION_LIST)
  //       .items.select("ConfigurationKey", "KeyValue")
  //       .filter(`ConfigurationKey eq '${configurationKey}'`)
  //       .top(1)
  //       .usingCaching()
  //       .get();

  //     return configurationValue && configurationValue.length > 0
  //       ? configurationValue[0].KeyValue
  //       : "";
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(`${strings.ConfigurationListErrorMessage}`);
  //   }
  // }

  // public async getImobilizadoById(
  //   id: string,
  //   empresa: string
  // ): Promise<IImobilizado[]> {
  //   try {
  //     const imobilizadoListId = await this.getConfigurationValue(
  //       IMOBILIZADO_CONF_KEY
  //     );
  //     if (imobilizadoListId) {
  //       const results: IImobilizado[] = await sp.web.lists
  //         .getById(imobilizadoListId)
  //         .items.select(
  //           "Id",
  //           "Classeimobilizado",
  //           "Descricao",
  //           "Imobilizado",
  //           "SubNumero",
  //           "ElementoPEP"
  //         )
  //         .filter(`Imobilizado eq '${id}' and Empresa eq ${empresa}`)
  //         .usingCaching()
  //         .orderBy("SubNumero")
  //         .get();
  //       return results && results.length > 0 ? results : ({} as IImobilizado[]);
  //     } else {
  //       const error: Error = new Error(strings.GetImobilizadoByIdErrorMessage);
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error.message);
  //   }
  // }

  // public async getMaterialById(id: string): Promise<IMaterial> {
  //   try {
  //     const materialListId = await this.getConfigurationValue(
  //       MATERIAL_CONF_KEY
  //     );
  //     if (materialListId) {
  //       const results: IMaterial[] = await sp.web.lists
  //         .getById(materialListId)
  //         .items.select("Id", "Material", "DescricaoItem")
  //         .filter(`Material eq '${id}'`)
  //         .top(1)
  //         .usingCaching()
  //         .get();
  //       return results && results.length > 0 ? results[0] : ({} as IMaterial);
  //     } else {
  //       const error: Error = new Error(`${strings.GetCentroByIdErrorMessage}`);
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error.message);
  //   }
  // }

  // public async getCentroById(id: string): Promise<ICentro> {
  //   try {
  //     const centroListId = await this.getConfigurationValue(CENTRO_CONF_KEY);
  //     if (centroListId) {
  //       const results: ICentro[] = await sp.web.lists
  //         .getById(centroListId)
  //         .items.select("Id", "Centro", "DescricaoItem")
  //         .filter(`Centro eq '${id}'`)
  //         .top(1)
  //         .usingCaching()
  //         .get();
  //       return results[0];
  //     } else {
  //       const error: Error = new Error(`${strings.GetCentroByIdErrorMessage}`);
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error.message);
  //   }
  // }

  // public async getGroupCompradoresById(
  //   id: string
  // ): Promise<IGrupoCompradores[]> {
  //   try {
  //     const EmpresaListId = await this.getConfigurationValue(
  //       GRUPOCOMPRADORES_CONF_KEY
  //     );
  //     if (EmpresaListId) {
  //       const results: IGrupoCompradores[] = await sp.web.lists
  //         .getById(EmpresaListId)
  //         .items.select("Id", "GrupoCompradores", "DescricaoItem")
  //         .filter(`GrupoCompradores eq '${id}'`)
  //         .top(1)
  //         .usingCaching()
  //         .get();
  //       return results;
  //     } else {
  //       const error: Error = new Error(
  //         `${strings.GetEmpresaByIdErrorMessage} `
  //       );
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error.message);
  //   }
  // }

  // public async getGroupMercadoriasById(
  //   id: string
  // ): Promise<IGrupoMercadorias[]> {
  //   try {
  //     const grupoMercadoriasListId = await this.getConfigurationValue(
  //       GRUPOMERCADORIAS_CONF_KEY
  //     );
  //     if (grupoMercadoriasListId) {
  //       const results: IGrupoMercadorias[] = await sp.web.lists
  //         .getById(grupoMercadoriasListId)
  //         .items.select("Id", "GrupoMercadorias", "DescricaoItem")
  //         .filter(`GrupoMercadorias eq '${id}'`)
  //         .top(1)
  //         .usingCaching()
  //         .get();
  //       return results;
  //     } else {
  //       const error: Error = new Error(
  //         `${strings.GetEmpresaByIdErrorMessage} `
  //       );
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error.message);
  //   }
  // }

  // /**
  //  * Gets empresa by id
  //  * @param id
  //  * @returns empresa by id
  //  */
  // public async getEmpresaById(id: string): Promise<IEmpresa> {
  //   try {
  //     const EmpresaListId = await this.getConfigurationValue(EMPRESA_CONF_KEY);
  //     if (EmpresaListId) {
  //       const results: IEmpresa[] = await sp.web.lists
  //         .getById(EmpresaListId)
  //         .items.select("Id", "Empresa", "DescricaoItem", "TipodePedido")
  //         .filter(`Empresa eq '${id}'`)
  //         .top(1)
  //         .usingCaching()
  //         .get();
  //       return results[0];
  //     } else {
  //       const error: Error = new Error(
  //         `${strings.GetEmpresaByIdErrorMessage} `
  //       );
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error.message);
  //   }
  // }

  // /**
  //  * Gets fornecedor by id
  //  * @param id
  //  * @returns fornecedor by id
  //  */
  // public async getFornecedorById(id: string): Promise<IFornecedor> {
  //   try {
  //     const fornecedorListId = await this.getConfigurationValue(
  //       FORNECEDOR_CONF_KEY
  //     );
  //     if (fornecedorListId) {
  //       const results: IFornecedor = await sp.web.lists
  //         .getById(fornecedorListId)
  //         .items.select("Id", "Fornecedor", "Descricao", "NIF")
  //         .filter(`Fornecedor eq '${id}'`)
  //         .top(1)
  //         .usingCaching()
  //         .get();
  //       return results;
  //     } else {
  //       const error: Error = new Error(strings.GetFornecedorByIdErrorMessage);
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error);
  //   }
  // }

  // public async getPaisById(id: string): Promise<IPais[]> {
  //   try {
  //     const paisListId = await this.getConfigurationValue(PAIS_CONF_KEY);
  //     if (paisListId) {
  //       const results: IPais[] = await sp.web.lists
  //         .getById(paisListId)
  //         .items.select("Id", "Pais", "DescricaoItem")
  //         .filter(`Pais eq '${id}'`)
  //         .top(1)
  //         .usingCaching()
  //         .get();
  //       return results;
  //     } else {
  //       const error: Error = new Error(
  //         "Lista de Países não existe na lista de configurações"
  //       );
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error);
  //   }
  // }

  // public async getOrdemById(id: string): Promise<IOrdem[]> {
  //   try {
  //     const paisListId = await this.getConfigurationValue(ORDEM_CONF_KEY);
  //     if (paisListId) {
  //       const results: IOrdem[] = await sp.web.lists
  //         .getById(paisListId)
  //         .items.select("Id", "Ordem", "DescricaoItem")
  //         .filter(`Ordem eq '${id}'`)
  //         .top(1)
  //         .usingCaching()
  //         .get();
  //       return results;
  //     } else {
  //       const error: Error = new Error(
  //         "Lista de Ordem não existe na lista de configurações"
  //       );
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error);
  //   }
  // }

  // public async getSegmentoMarcaById(id: string): Promise<ISegmentoMarca[]> {
  //   try {
  //     const segmentoMarcaListId = await this.getConfigurationValue(
  //       SEGMENTOMARCA_CONF_KEY
  //     );
  //     if (segmentoMarcaListId) {
  //       const results: ISegmentoMarca[] = await sp.web.lists
  //         .getById(segmentoMarcaListId)
  //         .items.select("Id", "SegmentoDeMarca", "DescricaoItem")
  //         .filter(`SegmentoDeMarca eq '${id}'`)
  //         .top(1)
  //         .usingCaching()
  //         .get();
  //       return results;
  //     } else {
  //       const error: Error = new Error(
  //         "Lista de Segmentos de marca não existe na lista de configurações"
  //       );
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error);
  //   }
  // }

  // public async getContaGastosById(id: string): Promise<IContaGastos[]> {
  //   try {
  //     const contaGastosListId = await this.getConfigurationValue(
  //       CONTAGASTOS_CONF_KEY
  //     );
  //     if (contaGastosListId) {
  //       const results: IContaGastos[] = await sp.web.lists
  //         .getById(contaGastosListId)
  //         .items.select("Id", "ContaGastos", "DescricaoItem")
  //         .filter(`ContaGastos eq '${id}'`)
  //         .top(1)
  //         .usingCaching()
  //         .get();
  //       return results;
  //     } else {
  //       const error: Error = new Error(
  //         "Lista de Conta Gastos não existe na lista de configurações"
  //       );
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error);
  //   }
  // }

  // /**
  //  * Gets canal by id
  //  * @param id
  //  * @returns canal by id
  //  */
  // public async getCanalById(id: string): Promise<ICanal[]> {
  //   try {
  //     const canalListId = await this.getConfigurationValue(CANAL_CONF_KEY);
  //     if (canalListId) {
  //       const results: ICanal[] = await sp.web.lists
  //         .getById(canalListId)
  //         .items.select("Id", "Canal", "DescricaoItem")
  //         .filter(`Canal eq '${id}'`)
  //         .top(1)
  //         .usingCaching()
  //         .get();
  //       return results;
  //     } else {
  //       const error: Error = new Error(strings.getCanalByIdErrorMessage);
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error);
  //   }
  // }

  // /**
  //  * Gets lista precos by id
  //  * @param id
  //  * @returns lista precos by id
  //  */
  // public async getListaPrecosById(id: string): Promise<IListaPrecos[]> {
  //   try {
  //     const listaPrecosListId = await this.getConfigurationValue(
  //       LISTAPRECOS_CONF_KEY
  //     );
  //     if (listaPrecosListId) {
  //       const results: IListaPrecos[] = await sp.web.lists
  //         .getById(listaPrecosListId)
  //         .items.select("Id", "ListaPrecos", "DescricaoItem")
  //         .filter(`ListaPrecos eq '${id}'`)
  //         .top(1)
  //         .usingCaching()
  //         .get();
  //       return results;
  //     } else {
  //       const error: Error = new Error(strings.getListaPrecosErrorMessage);
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error);
  //   }
  // }

  // /**
  //  * Gets sector by id
  //  * @param id
  //  * @returns sector by id
  //  */
  // public async getSectorById(id: string): Promise<ISector[]> {
  //   try {
  //     const sectotListId = await this.getConfigurationValue(SECTOR_CONF_KEY);
  //     if (sectotListId) {
  //       const results: ISector[] = await sp.web.lists
  //         .getById(sectotListId)
  //         .items.select("Id", "Sector", "DescricaoItem")
  //         .filter(`Sector eq '${id}'`)
  //         .top(1)
  //         .usingCaching()
  //         .get();
  //       return results;
  //     } else {
  //       const error: Error = new Error(strings.getSectorByIdErrorMessage);
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error);
  //   }
  // }

  // /**
  //  * Gets regiao vendas by id
  //  * @param id
  //  * @returns regiao vendas by id
  //  */
  // public async getRegiaoVendasById(id: string): Promise<IRegiaoVendas[]> {
  //   try {
  //     const regiaoVendasListId = await this.getConfigurationValue(
  //       REGIAO_VENDAS_CONF_KEY
  //     );
  //     if (regiaoVendasListId) {
  //       const results: IRegiaoVendas[] = await sp.web.lists
  //         .getById(regiaoVendasListId)
  //         .items.select("Id", "RegiaoVendas", "DescricaoItem")
  //         .filter(`RegiaoVendas eq '${id}'`)
  //         .top(1)
  //         .usingCaching()
  //         .get();
  //       return results;
  //     } else {
  //       const error: Error = new Error(strings.getRegiaoVendasByIdErrorMessage);
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error);
  //   }
  // }

  // public async getMarcaById(id: string): Promise<IMarca[]> {
  //   try {
  //     const regiaoVendasListId = await this.getConfigurationValue(
  //       MARCA_CONF_KEY
  //     );
  //     if (regiaoVendasListId) {
  //       const results: IMarca[] = await sp.web.lists
  //         .getById(regiaoVendasListId)
  //         .items.select("Id", "Marca", "DescricaoItem")
  //         .filter(`Marca eq '${id}'`)
  //         .top(1)
  //         .usingCaching()
  //         .get();
  //       return results;
  //     } else {
  //       const error: Error = new Error(strings.getRegiaoVendasByIdErrorMessage);
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error);
  //   }
  // }

  // public async getCentrosCustoById(id: string): Promise<ICentroCusto[]> {
  //   try {
  //     const centroCustosListId = await this.getConfigurationValue(
  //       CENTROCUSTOS_CONF_KEY
  //     );
  //     if (centroCustosListId) {
  //       const results: ICentroCusto[] = await sp.web.lists
  //         .getById(centroCustosListId)
  //         .items.select("Id", "CentroCustos", "DescricaoItem")
  //         .filter(`CentroCustos eq '${id}'`)
  //         .top(1)
  //         .usingCaching()
  //         .get();
  //       return results;
  //     } else {
  //       const error: Error = new Error(strings.getRegiaoVendasByIdErrorMessage);
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error);
  //   }
  // }

  // /**
  //  * Gets pedido compra by id
  //  * @param itemId
  //  * @returns pedido compra by id
  //  */
  // public async getPedidoCompraById(
  //   itemId: number
  // ): Promise<PagedItemCollection<IPedidoCompra[]>> {
  //   Log.verbose(
  //     LOG_SOURCE,
  //     "Start getPedidoCompraById..",
  //     this._context.serviceScope
  //   );
  //   const isApprover = await this.checkUserIsApprover();
  //   try {
  //     const pedidosCompraListId = await this.getConfigurationValue(
  //       PEDIDO_COMPRA_CONF_KEY
  //     );
  //     if (pedidosCompraListId) {
  //       // if current user is not approver don't get Item
  //       //  if (isApprover) {
  //       const results = await sp.web.lists
  //         .getById(pedidosCompraListId)
  //         .items.select(
  //           "Id",
  //           "Empresa",
  //           "Fornecedor",
  //           "NIF",
  //           "GrupoCompradores",
  //           "Solicitante",
  //           "DataPedido",
  //           "EstadoPedido",
  //           "ComentariosDoAprovador",
  //           "DescricaoEmpresa",
  //           "DescricaoFornecedor",
  //           "DescricaoGrupoComprador",
  //           "Observacoes",
  //           "Moeda",
  //           "Total",
  //           "Numero"
  //         )
  //         .filter(
  //           `Id eq '${itemId}'`
  //         )
  //         .top(10)
  //         .getPaged();
  //       Log.verbose(
  //         LOG_SOURCE,
  //         "End getPedidoCompraById..",
  //         this._context.serviceScope
  //       );
  //       return results;
  //       // } else {
  //       //   return undefined;
  //       // }
  //     } else {
  //       const error: Error = new Error(strings.GetPedidosCompraErrorMessage);
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error.message);
  //   }
  // }

  // public async getTotalPedidosCompraDetalhe(
  //   pedidoCompraId: string
  // ): Promise<number> {
  //   Log.verbose(
  //     LOG_SOURCE,
  //     "Start getPedidoCompraDetalhe..",
  //     this._context.serviceScope
  //   );

  //   let totalPedido: number = 0;
  //   try {
  //     const pedidosCompraDetalheListId = await this.getConfigurationValue(
  //       PEDIDO_COMPRA_DETALHE_CONF_KEY
  //     );

  //     if (pedidosCompraDetalheListId) {
  //       const results: IPedidoCompraDetalhe[] = await sp.web.lists
  //         .getById(pedidosCompraDetalheListId)
  //         .items.select("Total")
  //         .filter(`HeaderID eq '${pedidoCompraId}'`)
  //         .getAll(5000);

  //       for (const item of results) {
  //         totalPedido = totalPedido + item.Total;
  //       }
  //       return totalPedido;
  //     } else {
  //       const error: Error = new Error(strings.GetPedidosCompraErrorMessage);
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error.message);
  //   }
  // }

  // public async getPedidosCompraDetalhe(
  //   sortField: string,
  //   ascending: boolean,
  //   pedidoCompraId: string
  // ): Promise<PagedItemCollection<IPedidoCompra[]>> {
  //   Log.verbose(
  //     LOG_SOURCE,
  //     "Start getPedidoCompraDetalhe..",
  //     this._context.serviceScope
  //   );

  //   try {
  //     const pedidosCompraDetalheListId = await this.getConfigurationValue(
  //       PEDIDO_COMPRA_DETALHE_CONF_KEY
  //     );

  //     if (pedidosCompraDetalheListId) {
  //       const results = await sp.web.lists
  //         .getById(pedidosCompraDetalheListId)
  //         .items.select(
  //           "Id",
  //           "Centro",
  //           "Material",
  //           "DescricaoMaterial",
  //           "DescricaoCentro",
  //           "DescricaoItem",
  //           "ClassificacaoContabil",
  //           "Quantidade",
  //           "PrecoUnitario",
  //           "Total",
  //           "PorQuantidade",
  //           "UnidadeMedida"
  //         )
  //         .filter(`HeaderID eq '${pedidoCompraId}'`)
  //         .orderBy(`${sortField}`, ascending)
  //         .top(10)
  //         .getPaged();
  //       return results;
  //     } else {
  //       const error: Error = new Error(strings.GetPedidosCompraErrorMessage);
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error.message);
  //   }
  // }

  // // Get Pedido compra detalhe Item
  // public async getPedidosCompraDetalheById(
  //   id: string
  // ): Promise<IPedidoCompraDetalhe> {
  //   Log.verbose(
  //     LOG_SOURCE,
  //     "Start getPedidoCompraDetalhebyId..",
  //     this._context.serviceScope
  //   );

  //   try {
  //     const pedidosCompraDetalheListId = await this.getConfigurationValue(
  //       PEDIDO_COMPRA_DETALHE_CONF_KEY
  //     );

  //     if (pedidosCompraDetalheListId) {
  //       const results: IPedidoCompraDetalhe = await sp.web.lists
  //         .getById(pedidosCompraDetalheListId)
  //         .items.getById(Number(id))
  //         .get();
  //       return results;
  //     } else {
  //       const error: Error = new Error(strings.GetPedidosCompraErrorMessage);
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error.message);
  //   }
  // }

  // /**
  //  * Check user is approver
  //  */
  // private checkUserIsApprover = async (): Promise<Boolean> => {
  //   try {
  //     const AprovadorPedidosCompraListId = await this.getConfigurationValue(
  //       "AprovadorPedidoCompra"
  //     );
  //     if (AprovadorPedidosCompraListId) {
  //       const results = await sp.web.lists
  //         .getById(AprovadorPedidosCompraListId)
  //         .items.select("Id", "Aprovador/EMail", "Aprovador/Title")
  //         .filter(
  //           `Aprovador/EMail eq '${this._context.pageContext.user.email}'`
  //         )
  //         .expand("Aprovador")
  //         .top(1)
  //         .usingCaching()
  //         .get();
  //       return results.length > 0;
  //     } else {
  //       const error: Error = new Error(
  //         "Lista de Aprovadores Pedido Compra na defina na lista de configurações"
  //       );
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error.message);
  //   }
  // };

  // public searchPedidoCompraDetalhe = async (
  //   value: any,
  //   pedidoCompraId: string
  // ): Promise<PagedItemCollection<IPedidoCompra[]>> => {
  //   Log.verbose(
  //     LOG_SOURCE,
  //     "Start SerchPedidoCompra..",
  //     this._context.serviceScope
  //   );

  //   const searchQuantidade = isNaN(value.replace(",", "."))
  //     ? ""
  //     : ` or  (Quantidade eq ${value.replace(",", ".")})`;
  //   const searchPrecoUnitario = isNaN(value.replace(",", "."))
  //     ? ""
  //     : ` or  (PrecoUnitario eq ${value.replace(",", ".")})`;
  //   const searchTotal = isNaN(value.replace(",", "."))
  //     ? ""
  //     : ` or  (Total eq ${value.replace(",", ".")})`;

  //   const filterString = `HeaderID eq '${pedidoCompraId}'
  //    and ( substringof('${value}', Id)
  //     or substringof('${value}',Centro)
  //      or substringof('${value}',Material)
  //       or substringof('${value}',DescricaoItem)
  //        or substringof('${value}',DescricaoMaterial)
  //         or substringof('${value}',DescricaoCentro)
  //          or substringof('${value}',ClassificacaoContabil) ${searchPrecoUnitario} ${searchQuantidade} ${searchTotal})`;

  //   try {
  //     const pedidosCompraDetalheListId = await this.getConfigurationValue(
  //       PEDIDO_COMPRA_DETALHE_CONF_KEY
  //     );

  //     if (pedidosCompraDetalheListId) {
  //       const results = await sp.web.lists
  //         .getById(pedidosCompraDetalheListId)
  //         .items.select(
  //           "Id",
  //           "Centro",
  //           "UnidadeMedida",
  //           "Material",
  //           "DescricaoMaterial",
  //           "DescricaoCentro",
  //           "DescricaoItem",
  //           "ClassificacaoContabil",
  //           "Quantidade",
  //           "UnidadeMedida",
  //           "PrecoUnitario",
  //           "Total"
  //         )
  //         .filter(`${filterString}`)
  //         .orderBy(`Id`, false)
  //         .top(10)
  //         .getPaged();
  //       return results;

  //       // if current user is  approver show all items
  //     } else {
  //       const error: Error = new Error(
  //         strings.GetPedidosCompraDetalheErrorMessage
  //       );
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error.message);
  //   }
  // };

  // /**
  //  * Search pedido compra of spservices
  //  */
  // public searchPedidoCompra = async (
  //   value: any,
  //   sortColumn: string,
  //   sortDescending: boolean
  // ): Promise<PagedItemCollection<IPedidoCompra[]>> => {
  //   Log.verbose(
  //     LOG_SOURCE,
  //     "Start SerchPedidoCompra..",
  //     this._context.serviceScope
  //   );

  //   const searchTotal = isNaN(value.replace(",", "."))
  //     ? ""
  //     : ` or  (Total eq ${value.replace(",", ".")})`;

  //   value = encodeURIComponent(value.replace("'", ""));
  //   const isApprover = await this.checkUserIsApprover();
  //   try {
  //     const pedidosCompraListId = await this.getConfigurationValue(
  //       PEDIDO_COMPRA_CONF_KEY
  //     );
  //     if (pedidosCompraListId) {
  //       // if current user is not approver show only his items

  //       if (!isApprover) {
  //         let filterStringNotApprover = `Solicitante eq '${
  //           this._context.pageContext.user.email
  //         }' and (substringof('${value.replace(
  //           "'",
  //           "%27"
  //         )}', Id) or substringof( '${value}', Empresa) or substringof( '${value}', Fornecedor)  or substringof( '${value}',NIF) or  substringof( '${value}',EstadoPedido) or substringof( '${value}',DescricaoEmpresa)  or substringof( '${value}',DataPedido) or substringof( '${value}',DescricaoFornecedor) or substringof( '${value}',Numero) ${searchTotal})`;
  //         // let filterString = `(substringof('${value}', Id) or substringof('${value}', Empresa) or substringof('${value}', Fornecedor) or substringof('${value}',NIF) or substringof('${value}',EstadoPedido) or substringof('${value}',DescricaoEmpresa) or substringof('${value}',DescricaoFornecedor) or substringof('${value}',Numero) ${searchTotal})`;
  //         //  filterString =  `(Solicitante eq '${this._context.pageContext.user.email}') and ${filterString}`;
  //         const results = await sp.web.lists
  //           .getById(pedidosCompraListId)
  //           .items.select(
  //             "Id",
  //             "Empresa",
  //             "Fornecedor",
  //             "NIF",
  //             "GrupoCompradores",
  //             "Solicitante",
  //             "DataPedido",
  //             "EstadoPedido",
  //             "ComentariosDoAprovador",
  //             "DescricaoEmpresa",
  //             "DescricaoFornecedor",
  //             "Total",
  //             "Numero",
  //             "Moeda"
  //           )
  //           .filter(filterStringNotApprover)
  //           .orderBy(sortColumn, sortDescending)
  //           .top(10)
  //           .getPaged();
  //         return results;
  //       }
  //       // if current user is  approver show all items
  //       if (isApprover) {
  //         let filterStringIsApprover = `(substringof('${value}', Id) or substringof( '${value}', Empresa) or substringof( '${value}', Fornecedor)  or substringof( '${value}',NIF) or  substringof( '${value}',EstadoPedido) or substringof( '${value}',DescricaoEmpresa)  or substringof( '${value}',DataPedido) or substringof( '${value}',DescricaoFornecedor) or substringof( '${value}',Numero) ${searchTotal})`;

  //         const results = await sp.web.lists
  //           .getById(pedidosCompraListId)
  //           .items.select(
  //             "Id",
  //             "Empresa",
  //             "Fornecedor",
  //             "NIF",
  //             "GrupoCompradores",
  //             "Solicitante",
  //             "DataPedido",
  //             "EstadoPedido",
  //             "ComentariosDoAprovador",
  //             "DescricaoEmpresa",
  //             "DescricaoFornecedor",
  //             "Total",
  //             "Numero",
  //             "Moeda"
  //           )
  //           .orderBy(sortColumn, sortDescending)
  //           .top(14)
  //           .filter(filterStringIsApprover)
  //           .getPaged();
  //         Log.verbose(
  //           LOG_SOURCE,
  //           "End SeacrhPedidoCompra..",
  //           this._context.serviceScope
  //         );
  //         return results;
  //       }
  //     } else {
  //       const error: Error = new Error(strings.GetPedidosCompraErrorMessage);
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error.message);
  //   }
  // };

  // /**
  //  * Gets list items for list item picker
  //  * @param filterText
  //  * @param listId
  //  * @param internalColumnName
  //  * @param restApiUrl
  //  * @param [keyInternalColumnName]
  //  * @param [webUrl]
  //  * @param [filterList]
  //  * @param [restApiUrl]
  //  * @returns list items for list item picker
  //  */
  // public async getListItemsForListItemPicker(
  //   filterText: string,
  //   listId: string,
  //   internalColumnName: string,
  //   keyInternalColumnName?: string,
  //   webUrl?: string,
  //   restApiUrl?: string,
  //   filterList?: string
  // ): Promise<any[]> {
  //   let _filter: string = `$filter=startswith(${internalColumnName},'${encodeURIComponent(
  //     filterText.replace("'", "''")
  //   )}') `;
  //   let costumfilter: string = filterList ? `and ${filterList}` : "";
  //   let _top = " &$top=2000";

  //   // test wild character "*"  if "*" load first 30 items
  //   if (
  //     (filterText.trim().indexOf("*") == 0 && filterText.trim().length == 1) ||
  //     filterText.trim().length == 0
  //   ) {
  //     _filter = "";
  //     costumfilter = filterList ? `$filter=${filterList}&` : "";
  //     _top = `$top=500`;
  //   }
  //   try {
  //     let finalResponse;

  //     //TF 23 January this call do SP rest api will now get all items paginated and then return them concatenated
  //     let webAbsoluteUrl = !webUrl
  //       ? this._context.pageContext.web.absoluteUrl
  //       : webUrl;

  //     // If provided then use the provided url to make its own request, else the method will generate the rest api request
  //     let apiUrl = !restApiUrl
  //       ? `${webAbsoluteUrl}/_api/web/lists('${listId}')/items?$top=5000`
  //       : restApiUrl;

  //     console.log("Rest Api request ->" + apiUrl);

  //     let data = await this._context.spHttpClient.get(
  //       apiUrl,
  //       SPHttpClient.configurations.v1
  //     );
  //     if (data.ok) {
  //       // This will get data in batches of 5000 items while there still data to get
  //       let response = await data.json();

  //       finalResponse = response.value;

  //       // nextLink variable holds the link provided in the previous sharepoint rest api request, indicating that the result is paginated
  //       let nextLink = response["@odata.nextLink"];

  //       if (nextLink) {
  //         let recursiveResponse = await this.getListItemsForListItemPicker(
  //           filterText,
  //           listId,
  //           internalColumnName,
  //           keyInternalColumnName,
  //           webUrl,
  //           nextLink,
  //           filterList
  //         );
  //         console.log("After method execution");
  //         console.log(recursiveResponse);

  //         finalResponse = finalResponse.concat(recursiveResponse);
  //       }
  //     }

  //     return finalResponse;
  //   } catch (error) {
  //     return Promise.reject(error);
  //   }
  // }

  // public addPedidoCompraIem = async (
  //   pedidoCompraDetalhe: IPedidoCompraDetalhe
  // ): Promise<void> => {
  //   try {
  //     const pedidosCompraDetalheListId = await this.getConfigurationValue(
  //       PEDIDO_COMPRA_DETALHE_CONF_KEY
  //     );
  //     if (pedidosCompraDetalheListId) {
  //       const total: number =
  //         (pedidoCompraDetalhe.PrecoUnitario * pedidoCompraDetalhe.Quantidade) /
  //         pedidoCompraDetalhe.PorQuantidade;
  //       const results: ItemAddResult = await sp.web.lists
  //         .getById(pedidosCompraDetalheListId)
  //         .items.add({
  //           HeaderID: `${pedidoCompraDetalhe.HeaderID}`,
  //           Canal: `${
  //             pedidoCompraDetalhe.Canal ? pedidoCompraDetalhe.Canal : ""
  //           }`,
  //           Centro: `${
  //             pedidoCompraDetalhe.Centro ? pedidoCompraDetalhe.Centro : ""
  //           }`,
  //           CentroCustos: `${
  //             pedidoCompraDetalhe.CentroCustos
  //               ? pedidoCompraDetalhe.CentroCustos
  //               : ""
  //           }`,
  //           ClassificacaoContabil: `${pedidoCompraDetalhe.ClassificacaoContabil}`,
  //           ContaGastos: `${
  //             pedidoCompraDetalhe.ContaGastos
  //               ? pedidoCompraDetalhe.ContaGastos
  //               : ""
  //           }`,
  //           Deposito: `${
  //             pedidoCompraDetalhe.Deposito ? pedidoCompraDetalhe.Deposito : ""
  //           }`,
  //           DescricaoCentro: `${
  //             pedidoCompraDetalhe.DescricaoCentro
  //               ? pedidoCompraDetalhe.DescricaoCentro
  //               : ""
  //           }`,
  //           DescricaoCentroCustos: `${
  //             pedidoCompraDetalhe.DescricaoCentroCustos
  //               ? pedidoCompraDetalhe.DescricaoCentroCustos
  //               : ""
  //           }`,
  //           DescricaoDeposito: `${
  //             pedidoCompraDetalhe.DescricaoDeposito
  //               ? pedidoCompraDetalhe.DescricaoDeposito
  //               : ""
  //           }`,
  //           DescricaoCanal: `${
  //             pedidoCompraDetalhe.DescricaoCanal
  //               ? pedidoCompraDetalhe.DescricaoCanal
  //               : ""
  //           }`,
  //           DescricaoGrupoMercadorias: `${
  //             pedidoCompraDetalhe.DescricaoGrupoMercadorias
  //               ? pedidoCompraDetalhe.DescricaoGrupoMercadorias
  //               : ""
  //           }`,
  //           DescricaoImobilizado: `${
  //             pedidoCompraDetalhe.DescricaoImobilizado
  //               ? pedidoCompraDetalhe.DescricaoImobilizado
  //               : ""
  //           }`,
  //           DescricaoItem: `${
  //             pedidoCompraDetalhe.DescricaoItem
  //               ? pedidoCompraDetalhe.DescricaoItem
  //               : ""
  //           }`,
  //           DescricaoMarca: `${
  //             pedidoCompraDetalhe.DescricaoMarca
  //               ? pedidoCompraDetalhe.DescricaoMarca
  //               : ""
  //           }`,
  //           DescricaoMaterial: `${
  //             pedidoCompraDetalhe.DescricaoMaterial
  //               ? pedidoCompraDetalhe.DescricaoMaterial
  //               : ""
  //           }`,
  //           DescricaoOrdem: `${
  //             pedidoCompraDetalhe.DescricaoOrdem
  //               ? pedidoCompraDetalhe.DescricaoOrdem
  //               : ""
  //           }`,
  //           DescricaoContaGastos: `${
  //             pedidoCompraDetalhe.DescricaoContaGastos
  //               ? pedidoCompraDetalhe.DescricaoContaGastos
  //               : ""
  //           }`,
  //           DescricaoRegiaoVendas: `${
  //             pedidoCompraDetalhe.DescricaoRegiaoVendas
  //               ? pedidoCompraDetalhe.DescricaoRegiaoVendas
  //               : ""
  //           }`,
  //           DescricaoSegmentoMarca: `${
  //             pedidoCompraDetalhe.DescricaoSegmentoMarca
  //               ? pedidoCompraDetalhe.DescricaoSegmentoMarca
  //               : ""
  //           }`,
  //           DescricaoPais: `${
  //             pedidoCompraDetalhe.DescricaoPais
  //               ? pedidoCompraDetalhe.DescricaoPais
  //               : ""
  //           }`,
  //           DescricaoRubricaPandP: `${
  //             pedidoCompraDetalhe.DescricaoRubricaPandP
  //               ? pedidoCompraDetalhe.DescricaoRubricaPandP
  //               : ""
  //           }`,
  //           DescricaoListaPrecos: `${
  //             pedidoCompraDetalhe.DescricaoListaPrecos
  //               ? pedidoCompraDetalhe.DescricaoListaPrecos
  //               : ""
  //           }`,
  //           DescricaoSetor: `${
  //             pedidoCompraDetalhe.DescricaoSetor
  //               ? pedidoCompraDetalhe.DescricaoSetor
  //               : ""
  //           }`,
  //           GrupoMercadorias: `${
  //             pedidoCompraDetalhe.GrupoMercadorias
  //               ? pedidoCompraDetalhe.GrupoMercadorias
  //               : ""
  //           }`,
  //           Imobilizado: `${
  //             pedidoCompraDetalhe.Imobilizado
  //               ? pedidoCompraDetalhe.Imobilizado
  //               : ""
  //           }`,
  //           ListaPrecos: `${
  //             pedidoCompraDetalhe.ListaPrecos
  //               ? pedidoCompraDetalhe.ListaPrecos
  //               : ""
  //           }`,
  //           Marca: `${
  //             pedidoCompraDetalhe.Marca ? pedidoCompraDetalhe.Marca : ""
  //           }`,
  //           Material: `${
  //             pedidoCompraDetalhe.Material ? pedidoCompraDetalhe.Material : ""
  //           }`,
  //           Numero: `${
  //             pedidoCompraDetalhe.Numero ? pedidoCompraDetalhe.Numero : ""
  //           }`,
  //           Ordem: `${
  //             pedidoCompraDetalhe.Ordem ? pedidoCompraDetalhe.Ordem : ""
  //           }`,
  //           Pais: `${pedidoCompraDetalhe.Pais ? pedidoCompraDetalhe.Pais : ""}`,
  //           PorQuantidade: `${pedidoCompraDetalhe.PorQuantidade}`,
  //           PrecoUnitario: `${pedidoCompraDetalhe.PrecoUnitario}`,
  //           Quantidade: `${pedidoCompraDetalhe.Quantidade}`,
  //           RegiaoVendas: `${
  //             pedidoCompraDetalhe.RegiaoVendas
  //               ? pedidoCompraDetalhe.RegiaoVendas
  //               : ""
  //           }`,
  //           RubricaPandP: `${
  //             pedidoCompraDetalhe.RubricaPandP
  //               ? pedidoCompraDetalhe.RubricaPandP
  //               : ""
  //           }`,
  //           Sector: `${
  //             pedidoCompraDetalhe.Sector ? pedidoCompraDetalhe.Sector : ""
  //           }`,
  //           SegmentoDeMarca: `${
  //             pedidoCompraDetalhe.SegmentoDeMarca
  //               ? pedidoCompraDetalhe.SegmentoDeMarca
  //               : ""
  //           }`,
  //           SubNr: `${
  //             pedidoCompraDetalhe.SubNr ? pedidoCompraDetalhe.SubNr : ""
  //           }`,
  //           ElementoPEP: `${
  //             pedidoCompraDetalhe.ElementoPEP ? pedidoCompraDetalhe.ElementoPEP : ""
  //           }`,
  //           Total: total,
  //           UnidadeMedida: `${
  //             pedidoCompraDetalhe.UnidadeMedida
  //               ? pedidoCompraDetalhe.UnidadeMedida
  //               : ""
  //           }`,
  //           Moeda: `${pedidoCompraDetalhe.Moeda}`
  //         });
  //       const totalPedido = await this.getTotalPedidosCompraDetalhe(
  //         pedidoCompraDetalhe.HeaderID
  //       );
  //       await this.updateTotalPedidoCompra(
  //         Number(pedidoCompraDetalhe.HeaderID),
  //         totalPedido
  //       );
  //     } else {
  //       const error: Error = new Error(strings.GetPedidosCompraErrorMessage);
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error.message);
  //   }
  // };

  // public updatePedidoCompraIem = async (
  //   pedidoCompraDetalhe: IPedidoCompraDetalhe,
  //   id: number
  // ): Promise<void> => {
  //   try {
  //     const pedidosCompraDetalheListId = await this.getConfigurationValue(
  //       PEDIDO_COMPRA_DETALHE_CONF_KEY
  //     );
  //     //const pedidosCompraListId = await this.getConfigurationValue("Pedidocompra");
  //     if (pedidosCompraDetalheListId) {
  //       const total: number =
  //         (pedidoCompraDetalhe.PrecoUnitario * pedidoCompraDetalhe.Quantidade) /
  //         pedidoCompraDetalhe.PorQuantidade;
  //       const results: ItemUpdateResult = await sp.web.lists
  //         .getById(pedidosCompraDetalheListId)
  //         .items.getById(id)
  //         .update({
  //           HeaderID: `${pedidoCompraDetalhe.HeaderID}`,
  //           Canal: `${
  //             pedidoCompraDetalhe.Canal ? pedidoCompraDetalhe.Canal : ""
  //           }`,
  //           DescricaoCanal: `${
  //             pedidoCompraDetalhe.DescricaoCanal
  //               ? pedidoCompraDetalhe.DescricaoCanal
  //               : ""
  //           }`,
  //           Centro: `${
  //             pedidoCompraDetalhe.Centro ? pedidoCompraDetalhe.Centro : ""
  //           }`,
  //           CentroCustos: `${
  //             pedidoCompraDetalhe.CentroCustos
  //               ? pedidoCompraDetalhe.CentroCustos
  //               : ""
  //           }`,
  //           ClassificacaoContabil: `${pedidoCompraDetalhe.ClassificacaoContabil}`,
  //           ContaGastos: `${
  //             pedidoCompraDetalhe.ContaGastos
  //               ? pedidoCompraDetalhe.ContaGastos
  //               : ""
  //           }`,
  //           Deposito: `${
  //             pedidoCompraDetalhe.Deposito ? pedidoCompraDetalhe.Deposito : ""
  //           }`,
  //           DescricaoCentro: `${
  //             pedidoCompraDetalhe.DescricaoCentro
  //               ? pedidoCompraDetalhe.DescricaoCentro
  //               : ""
  //           }`,
  //           DescricaoCentroCustos: `${
  //             pedidoCompraDetalhe.DescricaoCentroCustos
  //               ? pedidoCompraDetalhe.DescricaoCentroCustos
  //               : ""
  //           }`,
  //           DescricaoDeposito: `${
  //             pedidoCompraDetalhe.DescricaoDeposito
  //               ? pedidoCompraDetalhe.DescricaoDeposito
  //               : ""
  //           }`,
  //           DescricaoGrupoMercadorias: `${
  //             pedidoCompraDetalhe.DescricaoGrupoMercadorias
  //               ? pedidoCompraDetalhe.DescricaoGrupoMercadorias
  //               : ""
  //           }`,
  //           DescricaoImobilizado: `${
  //             pedidoCompraDetalhe.DescricaoImobilizado
  //               ? pedidoCompraDetalhe.DescricaoImobilizado
  //               : ""
  //           }`,
  //           DescricaoItem: `${
  //             pedidoCompraDetalhe.DescricaoItem
  //               ? pedidoCompraDetalhe.DescricaoItem
  //               : ""
  //           }`,
  //           DescricaoMarca: `${
  //             pedidoCompraDetalhe.DescricaoMarca
  //               ? pedidoCompraDetalhe.DescricaoMarca
  //               : ""
  //           }`,
  //           DescricaoMaterial: `${
  //             pedidoCompraDetalhe.DescricaoMaterial
  //               ? pedidoCompraDetalhe.DescricaoMaterial
  //               : ""
  //           }`,
  //           DescricaoOrdem: `${
  //             pedidoCompraDetalhe.DescricaoOrdem
  //               ? pedidoCompraDetalhe.DescricaoOrdem
  //               : ""
  //           }`,
  //           DescricaoContaGastos: `${
  //             pedidoCompraDetalhe.DescricaoContaGastos
  //               ? pedidoCompraDetalhe.DescricaoContaGastos
  //               : ""
  //           }`,
  //           DescricaoSegmentoMarca: `${
  //             pedidoCompraDetalhe.DescricaoSegmentoMarca
  //               ? pedidoCompraDetalhe.DescricaoSegmentoMarca
  //               : ""
  //           }`,
  //           GrupoMercadorias: `${
  //             pedidoCompraDetalhe.GrupoMercadorias
  //               ? pedidoCompraDetalhe.GrupoMercadorias
  //               : ""
  //           }`,
  //           Imobilizado: `${
  //             pedidoCompraDetalhe.Imobilizado
  //               ? pedidoCompraDetalhe.Imobilizado
  //               : ""
  //           }`,
  //           ListaPrecos: `${
  //             pedidoCompraDetalhe.ListaPrecos
  //               ? pedidoCompraDetalhe.ListaPrecos
  //               : ""
  //           }`,
  //           DescricaoListaPrecos: `${
  //             pedidoCompraDetalhe.DescricaoListaPrecos
  //               ? pedidoCompraDetalhe.DescricaoListaPrecos
  //               : ""
  //           }`,
  //           Marca: `${
  //             pedidoCompraDetalhe.Marca ? pedidoCompraDetalhe.Marca : ""
  //           }`,
  //           Material: `${
  //             pedidoCompraDetalhe.Material ? pedidoCompraDetalhe.Material : ""
  //           }`,
  //           Numero: `${
  //             pedidoCompraDetalhe.Numero ? pedidoCompraDetalhe.Numero : ""
  //           }`,
  //           Ordem: `${
  //             pedidoCompraDetalhe.Ordem ? pedidoCompraDetalhe.Ordem : ""
  //           }`,
  //           Pais: `${pedidoCompraDetalhe.Pais ? pedidoCompraDetalhe.Pais : ""}`,
  //           DescricaoPais: `${
  //             pedidoCompraDetalhe.DescricaoPais
  //               ? pedidoCompraDetalhe.DescricaoPais
  //               : ""
  //           }`,
  //           PorQuantidade: `${pedidoCompraDetalhe.PorQuantidade}`,
  //           PrecoUnitario: `${pedidoCompraDetalhe.PrecoUnitario}`,
  //           Quantidade: `${pedidoCompraDetalhe.Quantidade}`,
  //           RegiaoVendas: `${
  //             pedidoCompraDetalhe.RegiaoVendas
  //               ? pedidoCompraDetalhe.RegiaoVendas
  //               : ""
  //           }`,
  //           DescricaoRegiaoVendas: `${
  //             pedidoCompraDetalhe.DescricaoRegiaoVendas
  //               ? pedidoCompraDetalhe.DescricaoRegiaoVendas
  //               : ""
  //           }`,
  //           RubricaPandP: `${
  //             pedidoCompraDetalhe.RubricaPandP
  //               ? pedidoCompraDetalhe.RubricaPandP
  //               : ""
  //           }`,
  //           DescricaoRubricaPandP: `${
  //             pedidoCompraDetalhe.DescricaoRubricaPandP
  //               ? pedidoCompraDetalhe.DescricaoRubricaPandP
  //               : ""
  //           }`,
  //           Sector: `${
  //             pedidoCompraDetalhe.Sector ? pedidoCompraDetalhe.Sector : ""
  //           }`,
  //           DescricaoSetor: `${
  //             pedidoCompraDetalhe.DescricaoSetor
  //               ? pedidoCompraDetalhe.DescricaoSetor
  //               : ""
  //           }`,
  //           SegmentoDeMarca: `${
  //             pedidoCompraDetalhe.SegmentoDeMarca
  //               ? pedidoCompraDetalhe.SegmentoDeMarca
  //               : ""
  //           }`,
  //           SubNr: `${
  //             pedidoCompraDetalhe.SubNr ? pedidoCompraDetalhe.SubNr : ""
  //           }`,
  //           ElementoPEP: `${
  //             pedidoCompraDetalhe.ElementoPEP ? pedidoCompraDetalhe.ElementoPEP : ""
  //           }`,
  //           Total: total,
  //           UnidadeMedida: `${
  //             pedidoCompraDetalhe.UnidadeMedida
  //               ? pedidoCompraDetalhe.UnidadeMedida
  //               : ""
  //           }`,
  //           Moeda: `${pedidoCompraDetalhe.Moeda}`
  //         });
  //       const totalPedido = await this.getTotalPedidosCompraDetalhe(
  //         pedidoCompraDetalhe.HeaderID
  //       );
  //       await this.updateTotalPedidoCompra(
  //         Number(pedidoCompraDetalhe.HeaderID),
  //         totalPedido
  //       );
  //     } else {
  //       const error: Error = new Error(strings.GetPedidosCompraErrorMessage);
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error.message);
  //   }
  // };

  // public deletePedidoCompraIem = async (
  //   headerId: string,
  //   id: number
  // ): Promise<void> => {
  //   try {
  //     const pedidosCompraDetalheListId = await this.getConfigurationValue(
  //       PEDIDO_COMPRA_DETALHE_CONF_KEY
  //     );
  //     //const pedidosCompraListId = await this.getConfigurationValue("Pedidocompra");
  //     if (pedidosCompraDetalheListId) {
  //       await sp.web.lists
  //         .getById(pedidosCompraDetalheListId)
  //         .items.getById(id)
  //         .delete();

  //       const totalPedido = await this.getTotalPedidosCompraDetalhe(headerId);
  //       await this.updateTotalPedidoCompra(Number(headerId), totalPedido);
  //     } else {
  //       const error: Error = new Error(strings.GetPedidosCompraErrorMessage);
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error.message);
  //   }
  // };

  // public async sentToAprove(listItemId: number) {
  //   try {
  //     const pedidosCompraListId = await this.getConfigurationValue(
  //       PEDIDO_COMPRA_CONF_KEY
  //     );

  //     const approvalFlowEndPoint = await this.getConfigurationValue(
  //       APPROVAL_FLOW_ENDPOINT_KEY
  //     );

  //     if (!approvalFlowEndPoint) {
  //       throw new Error(
  //         "Erro na submissão do Flow de aprovação, endereço do Flow não definido na lista de configurações."
  //       );
  //     }
  //     const httpClient: HttpClient = this._context.httpClient;
  //     //  const url: string = "https://prod-97.westeurope.logic.azure.com:443/workflows/a6fb7b73aa5b44fca81a4face39393eb/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=RkSj31yiLX5RF3_AHGVIy4IZTXG2R-iTTsHdJNJce4U";
  //     const url: string = approvalFlowEndPoint;
  //     const body: string = JSON.stringify({
  //       weburl: this._context.pageContext.site.absoluteUrl,
  //       listId: pedidosCompraListId,
  //       itemId: listItemId
  //     });
  //     const options: IHttpClientOptions = {
  //       body: body,
  //       headers: { "content-type": "application/json" }
  //     };
  //     const response: HttpClientResponse = await httpClient.post(
  //       url,
  //       HttpClient.configurations.v1,
  //       options
  //     );
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error.message);
  //   }
  // }

  // /**
  //  * Add pedido compra of spservices
  //  */
  // public addPedidoCompra = async (
  //   pedidoCompra: IPedidoCompra
  // ): Promise<void> => {
  //   try {
  //     const pedidosCompraListId = await this.getConfigurationValue(
  //       PEDIDO_COMPRA_CONF_KEY
  //     );
  //     //const pedidosCompraListId = await this.getConfigurationValue("Pedidocompra");
  //     if (pedidosCompraListId) {
  //       const results: ItemAddResult = await sp.web.lists
  //         .getById(pedidosCompraListId)
  //         .items.add({
  //           Empresa: `${pedidoCompra.Empresa}`,
  //           DescricaoEmpresa: `${pedidoCompra.DescricaoEmpresa}`,
  //           Fornecedor: `${pedidoCompra.Fornecedor}`,
  //           DescricaoFornecedor: `${pedidoCompra.DescricaoFornecedor}`,
  //           GrupoCompradores: `${pedidoCompra.GrupoCompradores}`,
  //           DescricaoGrupoComprador: `${pedidoCompra.DescricaoGrupoComprador}`,
  //           NIF: `${pedidoCompra.NIF}`,
  //           DataPedido: format(parseISO(new Date().toISOString()), "P", {
  //             locale: pt
  //           }),
  //           Solicitante: this._context.pageContext.user.email,
  //           EstadoPedido: estadoPedido.NaoSubmetido,
  //           Observacoes: pedidoCompra.Observacoes,
  //           Moeda: pedidoCompra.Moeda
  //         });
  //     } else {
  //       const error: Error = new Error(strings.GetPedidosCompraErrorMessage);
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error.message);
  //   }
  // };

  // /**
  //  * Update pedido compra of spservices
  //  */
  // public updatePedidoCompra = async (
  //   listItemId: number,
  //   pedidoCompra: IPedidoCompra
  // ): Promise<void> => {
  //   try {
  //     const pedidosCompraListId = await this.getConfigurationValue(
  //       PEDIDO_COMPRA_CONF_KEY
  //     );
  //     //const pedidosCompraListId = await this.getConfigurationValue("Pedidocompra");
  //     if (pedidosCompraListId) {
  //       const results: ItemAddResult = await sp.web.lists
  //         .getById(pedidosCompraListId)
  //         .items.getById(listItemId)
  //         .update({
  //           Empresa: `${pedidoCompra.Empresa}`,
  //           DescricaoEmpresa: `${pedidoCompra.DescricaoEmpresa}`,
  //           Fornecedor: `${pedidoCompra.Fornecedor}`,
  //           DescricaoFornecedor: `${pedidoCompra.DescricaoFornecedor}`,
  //           GrupoCompradores: `${pedidoCompra.GrupoCompradores}`,
  //           Observacoes: pedidoCompra.Observacoes,
  //           Moeda: pedidoCompra.Moeda,
  //           NIF: `${pedidoCompra.NIF}`
  //         });
  //     } else {
  //       const error: Error = new Error(strings.GetPedidosCompraErrorMessage);
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error.message);
  //   }
  // };

  // public updateTotalPedidoCompra = async (
  //   listItemId: number,
  //   total: number
  // ): Promise<void> => {
  //   try {
  //     const pedidosCompraListId = await this.getConfigurationValue(
  //       PEDIDO_COMPRA_CONF_KEY
  //     );

  //     if (pedidosCompraListId) {
  //       const results: ItemAddResult = await sp.web.lists
  //         .getById(pedidosCompraListId)
  //         .items.getById(listItemId)
  //         .update({
  //           Total: total
  //         });
  //     } else {
  //       const error: Error = new Error(strings.GetPedidosCompraErrorMessage);
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error.message);
  //   }
  // };

  // /**
  //  * Delete pedido compra of spservices
  //  */
  // public deletePedidoCompra = async (listItemId: number): Promise<void> => {
  //   try {
  //     const pedidosCompraDetalheListId = await this.getConfigurationValue(
  //       PEDIDO_COMPRA_DETALHE_CONF_KEY
  //     );
  //     const pedidosCompraListId = await this.getConfigurationValue(
  //       PEDIDO_COMPRA_CONF_KEY
  //     );

  //     // try to delete all details items
  //     if (pedidosCompraDetalheListId) {
  //       const results = await sp.web.lists
  //         .getById(pedidosCompraDetalheListId)
  //         .items.filter(`HeaderID eq '${listItemId}'`)
  //         .getAll();

  //       if (results && results.length > 0) {
  //         for (const listItem of results) {
  //           await sp.web.lists
  //             .getById(pedidosCompraDetalheListId)
  //             .items.getById(listItem.Id)
  //             .recycle();
  //         }
  //       }

  //       if (pedidosCompraListId) {
  //         await sp.web.lists
  //           .getById(pedidosCompraListId)
  //           .items.getById(listItemId)
  //           .recycle();
  //       } else {
  //         const error: Error = new Error(strings.GetPedidosCompraErrorMessage);
  //         Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //         throw new Error(error.message);
  //       }
  //     } else {
  //       const error: Error = new Error(strings.GetPedidosCompraErrorMessage);
  //       Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //       throw new Error(error.message);
  //     }
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error.message);
  //   }
  // };
  // /**
  //  * Gets user photo
  //  * @param userId Email or ID
  //  * @returns user photo
  //  */
  // public async getUserPhoto(userId): Promise<string> {
  //   const personaImgUrl = PROFILE_IMAGE_URL + userId;
  //   const url: string = await this.getImageBase64(personaImgUrl);
  //   const newHash = await this.getMd5HashForUrl(url);

  //   if (
  //     newHash !== DEFAULT_PERSONA_IMG_HASH &&
  //     newHash !== DEFAULT_IMAGE_PLACEHOLDER_HASH
  //   ) {
  //     return "data:image/png;base64," + url;
  //   } else {
  //     return "undefined";
  //   }
  // }

  // /**
  //  * Get MD5Hash for the image url to verify whether user has default image or custom image
  //  * @param url
  //  */
  // private getMd5HashForUrl(url: string) {
  //   return new Promise(async (resolve, reject) => {
  //     const library: any = await this.loadSPComponentById(MD5_MODULE_ID);
  //     try {
  //       const md5Hash = library.Md5Hash;
  //       if (md5Hash) {
  //         const convertedHash = md5Hash(url);
  //         resolve(convertedHash);
  //       }
  //     } catch (error) {
  //       resolve(url);
  //     }
  //   });
  // }

  // /**
  //  * Load SPFx component by id, SPComponentLoader is used to load the SPFx components
  //  * @param componentId - componentId, guid of the component library
  //  */
  // private loadSPComponentById(componentId: string) {
  //   return new Promise((resolve, reject) => {
  //     SPComponentLoader.loadComponentById(componentId)
  //       .then((component: any) => {
  //         resolve(component);
  //       })
  //       .catch(error => {});
  //   });
  // }
  // /**
  //  * Gets image base64
  //  * @param pictureUrl
  //  * @returns image base64
  //  */
  // private getImageBase64(pictureUrl: string): Promise<string> {
  //   return new Promise((resolve, reject) => {
  //     let image = new Image();
  //     image.addEventListener("load", () => {
  //       let tempCanvas = document.createElement("canvas");
  //       (tempCanvas.width = image.width),
  //         (tempCanvas.height = image.height),
  //         tempCanvas.getContext("2d").drawImage(image, 0, 0);
  //       let base64Str;
  //       try {
  //         base64Str = tempCanvas.toDataURL("image/png");
  //       } catch (e) {
  //         return "";
  //       }
  //       base64Str = base64Str.replace(/^data:image\/png;base64,/, "");
  //       resolve(base64Str);
  //     });
  //     image.src = pictureUrl;
  //   });
  // }

  // /**
  //  * Gets user profile
  //  * @param loginName
  //  * @returns user profile
  //  */
  // public async getUserProfile(loginName: string): Promise<any> {
  //   try {
  //     const _loginName = `i:0#.f|membership|${loginName}`;
  //     const user = await sp.profiles.getPropertiesFor(_loginName);
  //     /*
  //     const user = await graph.users
  //       .getById(loginName)
  //       .usingCaching()
  //       .get();*/
  //     return user;
  //   } catch (error) {
  //     Log.error(LOG_SOURCE, error, this._context.serviceScope);
  //     throw new Error(error.message);
  //   }
  // }
}

export declare class Log {
  private static _logHandler;
  /* Excluded from this release type: _initialize */
  /**
   * Logs a message which contains detailed information that is generally only needed for
   * troubleshooting.
   * @param   source - the source from where the message is logged, e.g., the class name.
   *          The source provides context information for the logged message.
   *          If the source's length is more than 20, only the first 20 characters are kept.
   * @param   message - the message to be logged
   *          If the message's length is more than 100, only the first 100 characters are kept.
   * @param   scope - the service scope that the source uses. A service scope can provide
    *         more context information (e.g., web part information) to the logged message.
   */
  static verbose(source: string, message: string, scope?: ServiceScope): void;
  /**
   * Logs a general informational message.
   * @param   source - the source from where the message is logged, e.g., the class name.
   *          The source provides context information for the logged message.
   *          If the source's length is more than 20, only the first 20 characters are kept.
   * @param   message - the message to be logged
   *          If the message's length is more than 100, only the first 100 characters are kept.
   * @param   scope - the service scope that the source uses. A service scope can provide
    *         more context information (e.g., web part information) to the logged message.
   */
  static info(source: string, message: string, scope?: ServiceScope): void;
  /**
   * Logs a warning.
   * @param   source - the source from where the message is logged, e.g., the class name.
   *          The source provides context information for the logged message.
   *          If the source's length is more than 20, only the first 20 characters are kept.
   * @param   message - the message to be logged
   *          If the message's length is more than 100, only the first 100 characters are kept.
   * @param   scope - the service scope that the source uses. A service scope can provide
    *         more context information (e.g., web part information) to the logged message.
   */
  static warn(source: string, message: string, scope?: ServiceScope): void;
  /**
   * Logs an error.
   * @param   source - the source from where the error is logged, e.g., the class name.
   *          The source provides context information for the logged error.
   *          If the source's length is more than 20, only the first 20 characters are kept.
   * @param   error - the error to be logged
   * @param   scope - the service scope that the source uses. A service scope can provide
    *         more context information (e.g., web part information) to the logged error.
   */
  static error(source: string, error: Error, scope?: ServiceScope): void;
}