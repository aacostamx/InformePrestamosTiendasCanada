use Catalogos
go

--Selecciona sólo las tiendas canada del catalogo catTiendas
--select NumeroTienda from CatTiendas where tipodetienda='C'

--Open data source que asegura que la tabla "CatTiendasMccPtmos" sean las más actuales
if exists(select * from sysobjects where name ='CatTiendasMccPtmos')drop table CatTiendasMccPtmos
select NumeroTienda, NumeroCiudad, TipoDeTienda,RegionTienda
into CatTiendasMccPtmos
from opendatasource('sqloledb', 'data source=10.0.17.123; user id=operacion; password=xoperacion').Catalogos.dbo.CatTiendas
where tipodetienda='C'